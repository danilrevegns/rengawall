# -*- coding: utf-8 -*-
"""
Создание пола и отделочных стен по контуру помещения в Renga (COM API).
Три пользовательских свойства помещения задают тип: по ним выбирается правило
(толщины, высота, стили, создавать ли пол/стены) из JSON-конфигурации.

Документация API: https://help.rengabim.com/api/
"""

from __future__ import annotations

import json
import math
import traceback
from typing import Any, Callable, Dict, Iterator, List, Optional, Tuple

# --- Константы сущностей и параметров (Renga API v2.46, строковые GUID) ---
ENTITY_ROOM = "{f1a805ff-573d-f46b-ffba-57f4bccaa6ed}"
ENTITY_WALL = "{4329112a-6b65-48d9-9da8-abf1f8f36327}"
ENTITY_FLOOR = "{f5bd8bd8-39c1-47f8-8499-f673c580dfbe}"

PARAM_WALL_HEIGHT = "{0c6c933c-e47c-40d2-ba84-b8ae5ccec6f1}"
PARAM_WALL_THICKNESS = "{25548335-7030-43b1-b602-9898f3adc3b0}"
PARAM_FLOOR_THICKNESS = "{f2712442-b9df-44fe-ac7b-c3524342c804}"

CURVE2D_UNDEFINED = 0
CURVE2D_LINE = 1
CURVE2D_ARC = 2
CURVE2D_POLY = 3

_MAX_POLY_RECURSE = 64

RECORD_PROG_ID = "Renga.Application.1"
DEFAULT_CONFIG_PATH = "room_finish_config.json"


def _norm_guid(g: str) -> str:
    g = g.strip().strip("{}").lower()
    return "{" + g + "}"


def _load_config(path: str) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    if "property_ids" not in data or len(data["property_ids"]) != 3:
        raise ValueError("В конфигурации нужен массив property_ids из трёх GUID.")
    data["property_ids"] = [_norm_guid(x) for x in data["property_ids"]]
    data.setdefault("defaults", {})
    data.setdefault("rules", [])
    return data


def _merge_rule(defaults: dict, rule: Optional[dict]) -> dict:
    out = dict(defaults)
    if rule:
        for k, v in rule.items():
            if k != "match":
                out[k] = v
    return out


def _find_rule(cfg: dict, triple: Tuple[str, str, str]) -> Optional[dict]:
    a, b, c = (x.strip().lower() for x in triple)

    def val(rule_match: dict, idx: int) -> str:
        return str(rule_match.get(str(idx), "")).strip().lower()

    for rule in cfg.get("rules", []):
        m = rule.get("match") or {}
        if val(m, 0) == a and val(m, 1) == b and val(m, 2) == c:
            return rule
    return None


def _get_property_text(prop) -> str:
    """Читает значение свойства в строку для сопоставления с правилами."""
    if prop is None:
        return ""
    try:
        if not prop.HasValue:
            return ""
    except Exception:
        pass
    t = getattr(prop, "Type", None)
    try:
        if t is not None:
            name = str(t)
            if "String" in name:
                return str(prop.GetStringValue()).strip()
    except Exception:
        pass
    for getter in (
        "GetStringValue",
        "GetEnumerationValue",
        "GetIntValue",
        "GetIntegerValue",
        "GetDoubleValue",
        "GetBoolValue",
    ):
        if hasattr(prop, getter):
            try:
                v = getattr(prop, getter)()
                if v is not None:
                    return str(v).strip()
            except Exception:
                continue
    return ""


def _room_type_triple(model_object, property_ids: List[str]) -> Tuple[str, str, str]:
    pc = model_object.GetProperties()
    out: List[str] = []
    for pid in property_ids:
        try:
            p = pc.GetS(pid)
        except Exception:
            p = None
        out.append(_get_property_text(p))
    return out[0], out[1], out[2]


def _load_renga_comtypes():
    """Типы Point2D/Placement3D из TLB Renga."""
    import comtypes.client.dynamic as dynamic
    import comtypes.gen.Renga as Renga

    return Renga, dynamic


def _dispatch_iface(ptr, dynamic_mod) -> Any:
    if ptr is None:
        return None
    try:
        return dynamic_mod.Dispatch(ptr)
    except Exception:
        return None


def _query_iface(com_obj, Renga, iface_name: str, dynamic_mod) -> Any:
    """
    Предпочтительно QueryInterface(типизированный интерфейс) — иначе Renga может
    падать с RPC_E_SERVERFAULT при вызовах через «ломаный» Dispatch.
    """
    iface = getattr(Renga, iface_name, None)
    if iface is not None:
        try:
            return com_obj.QueryInterface(iface)
        except Exception:
            pass
    return _dispatch_iface(com_obj.GetInterfaceByName(iface_name), dynamic_mod)


def _placement_axes_2d(Renga, placement) -> Tuple[Any, Any, Any]:
    """Origin, axis X, axis Y в плоскости этажа."""
    origin = placement.Origin
    axis_x = placement.AxisX
    try:
        axis_y = placement.AxisY
        if axis_y is not None:
            return origin, axis_x, axis_y
    except Exception:
        axis_y = None
    try:
        z = placement.AxisZ
        cx = float(z.Y) * float(axis_x.Z) - float(z.Z) * float(axis_x.Y)
        cy = float(z.Z) * float(axis_x.X) - float(z.X) * float(axis_x.Z)
        cz = float(z.X) * float(axis_x.Y) - float(z.Y) * float(axis_x.X)
        v = Renga.Vector3D()
        v.X, v.Y, v.Z = cx, cy, cz
        ln = math.hypot(v.X, v.Y) or 1.0
        v.X /= ln
        v.Y /= ln
        v.Z = 0.0
        return origin, axis_x, v
    except Exception:
        pass
    v = Renga.Vector3D()
    v.X = -float(axis_x.Y)
    v.Y = float(axis_x.X)
    v.Z = 0.0
    return origin, axis_x, v


def _sample_arc_points_xy(prim, n: int, log: Callable[[str], None]) -> List[Tuple[float, float]]:
    """Параметрическая выборка дуги (если API доступен), иначе только концы."""
    pts: List[Tuple[float, float]] = []
    for k in range(n + 1):
        t = k / n if n else 1.0
        p = None
        for name in (
            "Evaluate",
            "EvaluateCurve",
            "GetPoint",
            "GetPointByParameter",
            "GetParameterPoint",
        ):
            if not hasattr(prim, name):
                continue
            try:
                pt = getattr(prim, name)(t)
                p = (float(pt.X), float(pt.Y))
                break
            except Exception:
                continue
        if p is None:
            break
        if not pts or (
            abs(pts[-1][0] - p[0]) > 1e-9 or abs(pts[-1][1] - p[1]) > 1e-9
        ):
            pts.append(p)
    if len(pts) >= 2:
        return pts
    ends = _curve_endpoints(prim)
    if not ends:
        return []
    (sx, sy), (ex, ey) = ends
    return [(sx, sy), (ex, ey)]


def _outer_contour_vertex_chain_global(
    region_desc,
    Renga,
    log: Callable[[str], None],
) -> List[Tuple[float, float]]:
    """Упорядоченные вершины внешнего контура в ГСК (план)."""
    verts: List[Tuple[float, float]] = []
    for prim in _iter_outer_segments(region_desc, Renga):
        try:
            ct = int(prim.Curve2DType)
        except Exception:
            ct = CURVE2D_LINE
        if ct == CURVE2D_LINE:
            ends = _curve_endpoints(prim)
            if not ends:
                continue
            (sx, sy), (ex, ey) = ends
            if not verts:
                verts.append((sx, sy))
            verts.append((ex, ey))
        elif ct == CURVE2D_ARC:
            arc_pts = _sample_arc_points_xy(prim, 12, log)
            if not arc_pts:
                continue
            if not verts:
                verts.append(arc_pts[0])
            for p in arc_pts[1:]:
                verts.append(p)
        else:
            ends = _curve_endpoints(prim)
            if not ends:
                log(
                    "Сегмент типа %s без концов — пропуск при построении контура пола."
                    % ct
                )
                continue
            (sx, sy), (ex, ey) = ends
            if not verts:
                verts.append((sx, sy))
            verts.append((ex, ey))
    if len(verts) >= 2:
        a, b = verts[0], verts[-1]
        if math.hypot(a[0] - b[0], a[1] - b[1]) < 1e-3:
            verts.pop()
    return verts


def _vertex_chain_to_closed_global_lines(
    math_iface,
    Renga,
    verts: List[Tuple[float, float]],
    log: Callable[[str], None],
) -> List:
    """Замкнутый контур из отрезков в глобальных координатах."""
    if len(verts) < 2:
        return []
    cleaned: List[Tuple[float, float]] = [verts[0]]
    for p in verts[1:]:
        if (
            math.hypot(p[0] - cleaned[-1][0], p[1] - cleaned[-1][1]) > 1e-6
        ):
            cleaned.append(p)
    if len(cleaned) < 2:
        return []
    lines = []
    n = len(cleaned)
    for i in range(n):
        p1 = cleaned[i]
        p2 = cleaned[(i + 1) % n]
        try:
            lines.append(
                math_iface.CreateLineSegment2D(
                    _point2d(Renga, p1[0], p1[1]),
                    _point2d(Renga, p2[0], p2[1]),
                )
            )
        except Exception as ex:
            log("CreateLineSegment2D (контур пола): %s" % ex)
            return []
    return lines


def _local_xy_to_global_xy(
    lx: float,
    ly: float,
    origin,
    axis_x,
    axis_y,
) -> Tuple[float, float]:
    gx = float(origin.X) + lx * float(axis_x.X) + ly * float(axis_y.X)
    gy = float(origin.Y) + lx * float(axis_x.Y) + ly * float(axis_y.Y)
    return gx, gy


def _global_xy_to_local_xy(
    gx: float,
    gy: float,
    origin,
    axis_x,
    axis_y,
) -> Tuple[float, float]:
    """Обратное к test_floor_coor.to_global (только XY плана)."""
    vx = gx - float(origin.X)
    vy = gy - float(origin.Y)
    lx = vx * float(axis_x.X) + vy * float(axis_x.Y)
    ly = vx * float(axis_y.X) + vy * float(axis_y.Y)
    return lx, ly


def _curve_list_global_to_local_lines(
    math_iface,
    Renga,
    curves: List,
    origin,
    axis_x,
    axis_y,
    log: Callable[[str], None],
) -> List:
    """Только отрезки: перевод концов в ЛСК объекта перекрытия."""
    out: List = []
    for c in curves:
        try:
            ct = int(c.Curve2DType)
        except Exception:
            ct = CURVE2D_LINE
        if ct != CURVE2D_LINE:
            log("Пропуск сегмента не-линии при переводе контура пола в локальные координаты.")
            continue
        ends = _curve_endpoints(c)
        if not ends:
            continue
        (sx, sy), (ex, ey) = ends
        lsx, lsy = _global_xy_to_local_xy(sx, sy, origin, axis_x, axis_y)
        lex, ley = _global_xy_to_local_xy(ex, ey, origin, axis_x, axis_y)
        p1 = _point2d(Renga, lsx, lsy)
        p2 = _point2d(Renga, lex, ley)
        try:
            out.append(math_iface.CreateLineSegment2D(p1, p2))
        except Exception as ex:
            log("CreateLineSegment2D (local): %s" % ex)
    return out


def _ensure_closed_polyline_lines(
    math_iface,
    Renga,
    lines: List,
    tol: float = 1.0,
) -> List:
    """Замыкание контура перекрытия (мм), если последняя вершина не совпала с первой."""
    if len(lines) < 2:
        return lines
    try:
        c0 = lines[0]
        c1 = lines[-1]
        b = c0.GetBeginPoint()
        e = c1.GetEndPoint()
        dx = float(b.X) - float(e.X)
        dy = float(b.Y) - float(e.Y)
        if math.hypot(dx, dy) <= tol:
            return lines
        p1 = _point2d(Renga, float(e.X), float(e.Y))
        p2 = _point2d(Renga, float(b.X), float(b.Y))
        lines = list(lines)
        lines.append(math_iface.CreateLineSegment2D(p1, p2))
    except Exception:
        pass
    return lines


def _connect_app(prefer_running: bool):
    import comtypes.client

    if prefer_running:
        try:
            return comtypes.client.GetActiveObject(RECORD_PROG_ID)
        except Exception:
            pass
    return comtypes.client.CreateObject(RECORD_PROG_ID)


def _comtypes_renga_available() -> bool:
    try:
        _load_renga_comtypes()
        return True
    except Exception:
        return False


def _point2d(Renga, x: float, y: float):
    p = Renga.Point2D()
    p.X = float(x)
    p.Y = float(y)
    return p


def _point3d(Renga, x: float, y: float, z: float):
    p = Renga.Point3D()
    p.X = float(x)
    p.Y = float(y)
    p.Z = float(z)
    return p


def _vector3d(Renga, x: float, y: float, z: float):
    v = Renga.Vector3D()
    v.X = float(x)
    v.Y = float(y)
    v.Z = float(z)
    return v


def _placement2d_identity(Renga):
    pl = Renga.Placement2D()
    pl.Origin = _point2d(Renga, 0.0, 0.0)
    v = Renga.Vector2D()
    v.X, v.Y = 1.0, 0.0
    pl.xAxis = v
    return pl


def _curve_endpoints(curve) -> Optional[Tuple[Tuple[float, float], Tuple[float, float]]]:
    try:
        b = curve.GetBeginPoint()
        e = curve.GetEndPoint()
        return (float(b.X), float(b.Y)), (float(e.X), float(e.Y))
    except Exception:
        return None


def _atomic_curve_segments(curve, Renga, depth: int = 0) -> List:
    """
    Рекурсивно раскрывает контур до атомарных ICurve2D (отрезок/дуга).
    Сначала QueryInterface(IPolyCurve2D) — в TLB тип Curve2DType может быть Undefined,
    хотя контур — полилиния (документация Renga: cast к IPolyCurve2D).
    """
    if depth > _MAX_POLY_RECURSE:
        return [curve]
    poly = None
    try:
        poly = curve.QueryInterface(Renga.IPolyCurve2D)
    except Exception:
        poly = None
    if poly is not None:
        try:
            n = int(poly.GetSegmentCount())
        except Exception:
            n = 0
        if n <= 0:
            return [curve]
        out: List = []
        for i in range(n):
            try:
                sub = poly.GetSegment(int(i))
            except Exception:
                continue
            if sub is None:
                continue
            out.extend(_atomic_curve_segments(sub, Renga, depth + 1))
        return out if out else [curve]
    try:
        ct = int(curve.Curve2DType)
    except Exception:
        return [curve]
    if ct == CURVE2D_POLY:
        return [curve]
    return [curve]


def _ring_centroid_xy(verts: List[Tuple[float, float]]) -> Tuple[float, float]:
    if not verts:
        return 0.0, 0.0
    sx = sum(x for x, _ in verts)
    sy = sum(y for _, y in verts)
    n = float(len(verts))
    return sx / n, sy / n


def _inward_normal_for_segment(
    sx: float,
    sy: float,
    ex: float,
    ey: float,
    seed_x: float,
    seed_y: float,
    cx: float,
    cy: float,
) -> Optional[Tuple[float, float]]:
    """
    Единичная нормаль к отрезку, направленная к «внутренней» стороне (к опорной точке
    помещения или к центроиду контура), чтобы сместить baseline внутрь площади.
    """
    dx = ex - sx
    dy = ey - sy
    ln = math.hypot(dx, dy)
    if ln < 1e-9:
        return None
    ux, uy = dx / ln, dy / ln
    n1x, n1y = -uy, ux
    mx, my = (sx + ex) * 0.5, (sy + ey) * 0.5
    for px, py in ((seed_x, seed_y), (cx, cy)):
        vx, vy = px - mx, py - my
        dot = n1x * vx + n1y * vy
        if dot > 1e-6:
            return n1x, n1y
        if dot < -1e-6:
            return -n1x, -n1y
    return n1x, n1y


def _room_seed_point(room_iface) -> Optional[Tuple[float, float]]:
    try:
        if bool(room_iface.Automatic):
            cp = room_iface.ControlPoint
            return float(cp.X), float(cp.Y)
    except Exception:
        pass
    try:
        mp = room_iface.MarkerPosition
        return float(mp.X), float(mp.Y)
    except Exception:
        return None


def _try_calculate_room_region(
    room_iface, Renga, log: Callable[[str], None]
):
    """
    comtypes.gen.Renga.Point2D совместим с CalculateRegion (как в тестах).
    """
    candidates = []
    try:
        if bool(room_iface.Automatic):
            candidates.append(("ControlPoint", room_iface.ControlPoint))
    except Exception:
        pass
    try:
        candidates.append(("MarkerPosition", room_iface.MarkerPosition))
    except Exception:
        pass

    last_error: Optional[Exception] = None
    for _name, pt in candidates:
        if pt is None:
            continue
        try:
            rd = room_iface.CalculateRegion(pt)
            if rd:
                return rd
        except Exception as ex:
            last_error = ex
            continue

    seed = _room_seed_point(room_iface)
    if seed:
        pt = Renga.Point2D()
        pt.X = float(seed[0])
        pt.Y = float(seed[1])
        try:
            rd = room_iface.CalculateRegion(pt)
            if rd:
                return rd
        except Exception as ex:
            last_error = ex

    if last_error is not None:
        log("CalculateRegion не удался: %s" % last_error)
    else:
        log(
            "CalculateRegion вернул пусто (точка вне помещения, на границе "
            "или нет подходящей опорной точки)."
        )
    return None


def _iter_outer_segments(region_desc, Renga) -> Iterator:
    reg = region_desc.Region
    outer = reg.GetOuterContour()
    for prim in _atomic_curve_segments(outer, Renga, 0):
        yield prim


def _set_floor_baseline(
    floor_mo,
    composite_curve,
    Renga,
    dynamic_mod,
    log: Callable[[str], None],
) -> bool:
    """Перекрытие: контур в ЛСК объекта — IBaseline2DObject.SetBaseline (test_floor_coor.py)."""
    bl = _query_iface(floor_mo, Renga, "IBaseline2DObject", dynamic_mod)
    if not bl:
        log("IBaseline2DObject недоступен для пола.")
        return False
    try:
        bl.SetBaseline(composite_curve)
        return True
    except Exception as ex:
        log("SetBaseline пола: %s" % ex)
        return False


def _try_composite_curve(math_iface, curves: List) -> Any:
    if not curves:
        return None
    try:
        return math_iface.CreateCompositeCurve2D(curves)
    except Exception:
        pass
    try:
        return math_iface.CreateCompositeCurve2D(tuple(curves))
    except Exception:
        pass
    return None


def _set_params_double(obj, guid_s: str, value: float) -> None:
    params = obj.GetParameters()
    p = params.GetS(guid_s)
    if p:
        p.SetDoubleValue(float(value))


def _init_finish_wall_parameters(
    wall_mo, wh: float, wt: float
) -> None:
    """Параметры стены."""
    _set_params_double(wall_mo, PARAM_WALL_HEIGHT, wh)
    _set_params_double(wall_mo, PARAM_WALL_THICKNESS, wt)


def process_room(
    app,
    model,
    objects,
    room_mo,
    cfg: dict,
    log: Callable[[str], None],
) -> None:
    try:
        Renga, dynamic = _load_renga_comtypes()
    except Exception as ex:
        log(
            "Не загружен comtypes.gen.Renga (запустите скрипт один раз при открытой Renga "
            "или установите comtypes): %s" % ex
        )
        return

    if room_mo.ObjectTypeS.lower() != ENTITY_ROOM.lower():
        log("Пропуск: объект не помещение (id=%s)." % room_mo.Id)
        return

    room = _query_iface(room_mo, Renga, "IRoom", dynamic)
    if not room:
        log("Не удалось получить IRoom (id=%s)." % room_mo.Id)
        return

    triple = _room_type_triple(room_mo, cfg["property_ids"])
    rule = _find_rule(cfg, triple)
    merged = _merge_rule(cfg["defaults"], rule)
    if rule is None:
        log(
            "Помещение id=%s: нет правила для (%r, %r, %r), используются defaults."
            % (room_mo.Id, triple[0], triple[1], triple[2])
        )
    else:
        log(
            "Помещение id=%s: правило для (%r, %r, %r)."
            % (room_mo.Id, triple[0], triple[1], triple[2])
        )

    seed = _room_seed_point(room)
    if not seed:
        log("Не удалось опорную точку помещения id=%s." % room_mo.Id)
        return

    region_desc = _try_calculate_room_region(room, Renga, log)
    if not region_desc:
        return

    level_obj = _query_iface(room_mo, Renga, "ILevelObject", dynamic)
    if not level_obj:
        log("Нет ILevelObject у помещения.")
        return
    level_id = int(level_obj.LevelId)
    level_mo = objects.GetById(level_id)
    ilvl = _query_iface(level_mo, Renga, "ILevel", dynamic)
    if not ilvl:
        log("Не найден уровень id=%s." % level_id)
        return

    math = app.Math
    project = app.Project

    floor_verts = _outer_contour_vertex_chain_global(region_desc, Renga, log)
    floor_global_lines = _vertex_chain_to_closed_global_lines(
        math, Renga, floor_verts, log
    )
    if not floor_global_lines:
        log("Нет сегментов внешнего контура помещения для пола.")
        return

    if not merged.get("create_floor", True) and not merged.get("create_walls", True):
        log("Помещение id=%s: create_floor и create_walls выключены — пропуск." % room_mo.Id)
        return

    op = project.CreateOperation()
    op.Start()
    try:
        if merged.get("create_floor", True):
            args = model.CreateNewEntityArgs()
            args.TypeIdS = ENTITY_FLOOR
            args.HostObjectId = level_id
            sid = int(merged.get("floor_style_id", 0) or 0)
            if sid:
                args.StyleId = sid
            floor_mo = model.CreateObject(args)
            if floor_mo:
                ft = float(merged.get("floor_thickness_mm", 50))
                _set_params_double(floor_mo, PARAM_FLOOR_THICKNESS, ft)
                pl_floor = None
                disp_fl = _query_iface(
                    floor_mo, Renga, "ILevelObject", dynamic
                )
                if disp_fl:
                    try:
                        pl_floor = disp_fl.Placement
                    except Exception:
                        try:
                            pl_floor = disp_fl.GetPlacement()
                        except Exception:
                            pl_floor = None
                if not pl_floor:
                    try:
                        pl_floor = ilvl.Placement
                    except Exception:
                        try:
                            pl_floor = ilvl.GetPlacement()
                        except Exception:
                            pl_floor = None
                comp_floor = None
                if pl_floor and floor_global_lines:
                    o, ax, ay = _placement_axes_2d(Renga, pl_floor)
                    loc_lines = _curve_list_global_to_local_lines(
                        math, Renga, floor_global_lines, o, ax, ay, log
                    )
                    loc_lines = _ensure_closed_polyline_lines(
                        math, Renga, loc_lines
                    )
                    comp_floor = _try_composite_curve(math, loc_lines)
                if not comp_floor:
                    log(
                        "Контур пола в ЛСК не собран — пробую глобальный контур "
                        "(при сбое Renga проверьте контур и Placement перекрытия)."
                    )
                    comp_floor = _try_composite_curve(math, floor_global_lines)
                if comp_floor and _set_floor_baseline(
                    floor_mo, comp_floor, Renga, dynamic, log
                ):
                    log("Пол создан (id=%s)." % floor_mo.Id)
                else:
                    log(
                        "Пол создан (id=%s), контур не применён — проверьте модель."
                        % floor_mo.Id
                    )
            else:
                log("CreateObject(Floor) не вернул объект.")

        if merged.get("create_walls", True):
            wh = float(merged.get("wall_height_mm", 3000))
            wt = float(merged.get("wall_thickness_mm", 120))
            wstyle = int(merged.get("wall_style_id", 0) or 0)
            
            cx_c, cy_c = _ring_centroid_xy(floor_verts)
            geom_inset = merged.get("wall_geometric_inset_from_contour", True)
            raw_inset = merged.get("wall_contour_inset_mm", None)
            if not geom_inset:
                wall_inset_mm = 0.0
            elif raw_inset is not None:
                wall_inset_mm = float(raw_inset)
            else:
                wall_inset_mm = wt * 0.5
                
            for prim in _iter_outer_segments(region_desc, Renga):
                wall_mo = None
                try:
                    ct = int(prim.Curve2DType)
                except Exception:
                    ct = CURVE2D_LINE
                args = model.CreateNewEntityArgs()
                args.TypeIdS = ENTITY_WALL
                args.HostObjectId = level_id
                if wstyle:
                    args.StyleId = wstyle
                if ct == CURVE2D_LINE:
                    ends = _curve_endpoints(prim)
                    if not ends:
                        continue
                    (sx, sy), (ex, ey) = ends
                    if wall_inset_mm > 1e-6:
                        nrm = _inward_normal_for_segment(
                            sx,
                            sy,
                            ex,
                            ey,
                            seed[0],
                            seed[1],
                            cx_c,
                            cy_c,
                        )
                        if nrm is not None:
                            nx, ny = nrm
                            sx += nx * wall_inset_mm
                            sy += ny * wall_inset_mm
                            ex += nx * wall_inset_mm
                            ey += ny * wall_inset_mm
                    
                    wall_mo = model.CreateObject(args)
                    if not wall_mo:
                        continue
                    _init_finish_wall_parameters(wall_mo, wh, wt)
                    bl = _query_iface(
                        wall_mo, Renga, "IBaseline2DObject", dynamic
                    )
                    if bl:
                        loc = math.CreateLineSegment2D(
                            _point2d(Renga, sx, sy),
                            _point2d(Renga, ex, ey),
                        )
                        bl.SetBaseline(loc)
                elif ct == CURVE2D_ARC:
                    wall_mo = model.CreateObject(args)
                    if not wall_mo:
                        continue
                    _init_finish_wall_parameters(wall_mo, wh, wt)
                    bl = _query_iface(
                        wall_mo, Renga, "IBaseline2DObject", dynamic
                    )
                    if bl:
                        try:
                            arc_c = prim.GetCopy()
                            bl.SetBaselineInCS(
                                _placement2d_identity(Renga), arc_c
                            )
                        except Exception as ex:
                            log("Дуга стены id=%s: %s" % (wall_mo.Id, ex))
                else:
                    log(
                        "Пропуск сегмента неподдерживаемого типа кривой: %s" % ct
                    )
                    continue
        op.Apply()
        log("Изменения применены для помещения id=%s." % room_mo.Id)
    except Exception:
        op.Rollback()
        log("Откат операции: %s" % traceback.format_exc())
        raise


def collect_room_ids(
    app, model, mode: str, explicit: List[int]
) -> List[int]:
    objects = model.GetObjects()
    if mode == "selection":
        sel = app.Selection
        ids = list(sel.GetSelectedObjects())
        rooms = []
        for oid in ids:
            mo = objects.GetById(int(oid))
            if mo and mo.ObjectTypeS.lower() == ENTITY_ROOM.lower():
                rooms.append(int(oid))
        return rooms
    if mode == "all":
        rooms = []
        try:
            for obj_id in objects.GetIds():
                mo = objects.GetById(int(obj_id))
                if mo and mo.ObjectTypeS.lower() == ENTITY_ROOM.lower():
                    rooms.append(int(mo.Id))
        except Exception:
            n = int(objects.Count)
            for i in range(n):
                mo = objects.GetByIndex(i)
                if mo and mo.ObjectTypeS.lower() == ENTITY_ROOM.lower():
                    rooms.append(int(mo.Id))
        return rooms
    return list(explicit)


def run_batch(
    config_path: str,
    room_ids: List[int],
    mode: str,
    prefer_running: bool,
    log: Callable[[str], None],
) -> None:
    if not _comtypes_renga_available():
        raise RuntimeError(
            "Не импортируется comtypes.gen.Renga. Выполните: pip install comtypes, "
            "запустите Renga с проектом и повторите (typelib генерируется при первом доступе)."
        )
    cfg = _load_config(config_path)
    app = _connect_app(prefer_running)
    app.Visible = True
    if not app.HasProject:
        raise RuntimeError("В Renga нет открытого проекта.")
    project = app.Project
    model = project.Model
    objects = model.GetObjects()
    ids = collect_room_ids(app, model, mode, room_ids)
    if not ids:
        raise RuntimeError("Нет помещений для обработки (проверьте выбор или список id).")
    for rid in ids:
        mo = objects.GetById(int(rid))
        if not mo:
            log("Объект id=%s не найден." % rid)
            continue
        process_room(app, model, objects, mo, cfg, log)
