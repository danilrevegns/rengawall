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


def norm_guid(g: str) -> str:
    g = g.strip().strip("{}").lower()
    return "{" + g + "}"


def load_cfg(path: str) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    if "property_ids" not in data or len(data["property_ids"]) != 3:
        raise ValueError("В конфигурации нужен массив property_ids из трёх GUID.")
    data["property_ids"] = [norm_guid(x) for x in data["property_ids"]]
    data.setdefault("defaults", {})
    data.setdefault("rules", [])
    return data


def merge_rule(defaults: dict, rule: Optional[dict]) -> dict:
    out = dict(defaults)
    if rule:
        for k, v in rule.items():
            if k != "match":
                out[k] = v
    return out


def find_rule(cfg: dict, triple: Tuple[str, str, str]) -> Optional[dict]:
    a, b, c = (x.strip().lower() for x in triple)
    for rule in cfg.get("rules", []):
        m = rule.get("match") or {}
        if str(m.get("0", "")).strip().lower() == a and \
           str(m.get("1", "")).strip().lower() == b and \
           str(m.get("2", "")).strip().lower() == c:
            return rule
    return None


def get_prop(prop) -> str:
    """Читает значение свойства в строку для сопоставления с правилами."""
    if not prop: return ""
    try:
        if not prop.HasValue: return ""
    except Exception: pass

    if "String" in str(getattr(prop, "Type", "")):
        try: return str(prop.GetStringValue()).strip()
        except Exception: pass

    for getter in ("GetStringValue", "GetEnumerationValue", "GetIntValue", "GetIntegerValue", "GetDoubleValue", "GetBoolValue"):
        if hasattr(prop, getter):
            try:
                v = getattr(prop, getter)()
                if v is not None: return str(v).strip()
            except Exception: pass
    return ""


def room_type(model_object, property_ids: List[str]) -> Tuple[str, str, str]:
    pc = model_object.GetProperties()
    def get_p(pid):
        try: return get_prop(pc.GetS(pid))
        except Exception: return ""
    res = [get_p(pid) for pid in property_ids[:3]]
    return res[0], res[1], res[2]


def load_com():
    """Типы Point2D/Placement3D из TLB Renga."""
    import comtypes.client.dynamic as dynamic
    import comtypes.gen.Renga as Renga

    return Renga, dynamic


def dispatch(ptr, dynamic_mod) -> Any:
    if ptr is None:
        return None
    try:
        return dynamic_mod.Dispatch(ptr)
    except Exception:
        return None


def query(com_obj, Renga, iface_name: str, dynamic_mod) -> Any:
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
    return dispatch(com_obj.GetInterfaceByName(iface_name), dynamic_mod)


def get_axes(Renga, placement) -> Tuple[Any, Any, Any]:
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


def sample_arc(prim, n: int, log: Callable[[str], None]) -> List[Tuple[float, float]]:
    pts: List[Tuple[float, float]] = []
    methods = ("Evaluate", "EvaluateCurve", "GetPoint", "GetPointByParameter", "GetParameterPoint")
    for k in range(n + 1):
        t = k / n if n else 1.0
        p = None
        for name in methods:
            if hasattr(prim, name):
                try:
                    pt = getattr(prim, name)(t)
                    p = (float(pt.X), float(pt.Y))
                    break
                except Exception: pass
        if not p: break
        if not pts or abs(pts[-1][0] - p[0]) > 1e-9 or abs(pts[-1][1] - p[1]) > 1e-9:
            pts.append(p)
    if len(pts) >= 2: return pts
    ends = curve_ends(prim)
    return list(ends) if ends else []


def get_verts(region_desc, Renga, log: Callable[[str], None]) -> List[Tuple[float, float]]:
    verts: List[Tuple[float, float]] = []
    for prim in iter_outer(region_desc, Renga):
        ct = getattr(prim, "Curve2DType", CURVE2D_LINE)
        if ct == CURVE2D_ARC:
            pts = sample_arc(prim, 12, log)
        else:
            pts = list(curve_ends(prim) or [])
            if not pts and ct != CURVE2D_LINE:
                log(f"Сегмент {ct} без концов — пропуск.")
                
        for p in pts:
            if not verts or math.hypot(verts[-1][0] - p[0], verts[-1][1] - p[1]) > 1e-6:
                verts.append(p)
                
    if len(verts) >= 2 and math.hypot(verts[0][0] - verts[-1][0], verts[0][1] - verts[-1][1]) < 1e-3:
        verts.pop()
    return verts


def get_lines(math_iface, Renga, verts: List[Tuple[float, float]], log: Callable[[str], None]) -> List:
    if len(verts) < 2: return []
    lines = []
    for i, p1 in enumerate(verts):
        p2 = verts[(i + 1) % len(verts)]
        try: lines.append(math_iface.CreateLineSegment2D(pt2d(Renga, *p1), pt2d(Renga, *p2)))
        except Exception as ex:
            log(f"CreateLineSegment2D: {ex}")
            return []
    return lines


def to_global(
    lx: float,
    ly: float,
    origin,
    axis_x,
    axis_y,
) -> Tuple[float, float]:
    gx = float(origin.X) + lx * float(axis_x.X) + ly * float(axis_y.X)
    gy = float(origin.Y) + lx * float(axis_x.Y) + ly * float(axis_y.Y)
    return gx, gy


def to_local(
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


def local_lines(
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
        ends = curve_ends(c)
        if not ends:
            continue
        (sx, sy), (ex, ey) = ends
        lsx, lsy = to_local(sx, sy, origin, axis_x, axis_y)
        lex, ley = to_local(ex, ey, origin, axis_x, axis_y)
        p1 = pt2d(Renga, lsx, lsy)
        p2 = pt2d(Renga, lex, ley)
        try:
            out.append(math_iface.CreateLineSegment2D(p1, p2))
        except Exception as ex:
            log("CreateLineSegment2D (local): %s" % ex)
    return out


def close_lines(
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
        p1 = pt2d(Renga, float(e.X), float(e.Y))
        p2 = pt2d(Renga, float(b.X), float(b.Y))
        lines = list(lines)
        lines.append(math_iface.CreateLineSegment2D(p1, p2))
    except Exception:
        pass
    return lines


def connect(prefer_running: bool):
    import comtypes.client

    if prefer_running:
        try:
            return comtypes.client.GetActiveObject(RECORD_PROG_ID)
        except Exception:
            pass
    return comtypes.client.CreateObject(RECORD_PROG_ID)


def has_renga() -> bool:
    try:
        load_com()
        return True
    except Exception:
        return False


def pt2d(Renga, x: float, y: float):
    p = Renga.Point2D()
    p.X = float(x)
    p.Y = float(y)
    return p


def pt3d(Renga, x: float, y: float, z: float):
    p = Renga.Point3D()
    p.X = float(x)
    p.Y = float(y)
    p.Z = float(z)
    return p


def vec3d(Renga, x: float, y: float, z: float):
    v = Renga.Vector3D()
    v.X = float(x)
    v.Y = float(y)
    v.Z = float(z)
    return v


def ident_pl(Renga):
    pl = Renga.Placement2D()
    pl.Origin = pt2d(Renga, 0.0, 0.0)
    v = Renga.Vector2D()
    v.X, v.Y = 1.0, 0.0
    pl.xAxis = v
    return pl


def curve_ends(curve) -> Optional[Tuple[Tuple[float, float], Tuple[float, float]]]:
    try:
        b = curve.GetBeginPoint()
        e = curve.GetEndPoint()
        return (float(b.X), float(b.Y)), (float(e.X), float(e.Y))
    except Exception:
        return None


def get_segs(curve, Renga, depth: int = 0) -> List:
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
            out.extend(get_segs(sub, Renga, depth + 1))
        return out if out else [curve]
    try:
        ct = int(curve.Curve2DType)
    except Exception:
        return [curve]
    if ct == CURVE2D_POLY:
        return [curve]
    return [curve]


def centroid(verts: List[Tuple[float, float]]) -> Tuple[float, float]:
    if not verts:
        return 0.0, 0.0
    sx = sum(x for x, _ in verts)
    sy = sum(y for _, y in verts)
    n = float(len(verts))
    return sx / n, sy / n


def inward_norm(
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


def room_seed(room_iface) -> Optional[Tuple[float, float]]:
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


def calc_region(room_iface, Renga, log: Callable[[str], None]):
    pts = []
    if getattr(room_iface, "Automatic", False):
        pts.append(getattr(room_iface, "ControlPoint", None))
    pts.append(getattr(room_iface, "MarkerPosition", None))
    
    seed = room_seed(room_iface)
    if seed: pts.append(pt2d(Renga, *seed))
    
    last_err = None
    for pt in filter(None, pts):
        try:
            if rd := room_iface.CalculateRegion(pt): return rd
        except Exception as ex: last_err = ex
        
    log(f"CalculateRegion не удался: {last_err}" if last_err else "CalculateRegion вернул пусто.")
    return None


def iter_outer(region_desc, Renga) -> Iterator:
    reg = region_desc.Region
    outer = reg.GetOuterContour()
    for prim in get_segs(outer, Renga, 0):
        yield prim


def set_base(
    floor_mo,
    composite_curve,
    Renga,
    dynamic_mod,
    log: Callable[[str], None],
) -> bool:
    """Перекрытие: контур в ЛСК объекта — IBaseline2DObject.SetBaseline (test_floor_coor.py)."""
    bl = query(floor_mo, Renga, "IBaseline2DObject", dynamic_mod)
    if not bl:
        log("IBaseline2DObject недоступен для пола.")
        return False
    try:
        bl.SetBaseline(composite_curve)
        return True
    except Exception as ex:
        log("SetBaseline пола: %s" % ex)
        return False


def comp_curve(math_iface, curves: List) -> Any:
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


def set_param(obj, guid_s: str, value: float) -> None:
    params = obj.GetParameters()
    p = params.GetS(guid_s)
    if p:
        p.SetDoubleValue(float(value))


def init_wall(
    wall_mo, wh: float, wt: float
) -> None:
    """Параметры стены."""
    set_param(wall_mo, PARAM_WALL_HEIGHT, wh)
    set_param(wall_mo, PARAM_WALL_THICKNESS, wt)


def proc_room(
    app,
    model,
    objects,
    room_mo,
    cfg: dict,
    log: Callable[[str], None],
) -> None:
    try:
        Renga, dynamic = load_com()
    except Exception as ex:
        log(
            "Не загружен comtypes.gen.Renga (запустите скрипт один раз при открытой Renga "
            "или установите comtypes): %s" % ex
        )
        return

    if room_mo.ObjectTypeS.lower() != ENTITY_ROOM.lower():
        log("Пропуск: объект не помещение (id=%s)." % room_mo.Id)
        return

    room = query(room_mo, Renga, "IRoom", dynamic)
    if not room:
        log("Не удалось получить IRoom (id=%s)." % room_mo.Id)
        return

    triple = room_type(room_mo, cfg["property_ids"])
    rule = find_rule(cfg, triple)
    merged = merge_rule(cfg["defaults"], rule)
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

    seed = room_seed(room)
    if not seed:
        log("Не удалось опорную точку помещения id=%s." % room_mo.Id)
        return

    region_desc = calc_region(room, Renga, log)
    if not region_desc:
        return

    level_obj = query(room_mo, Renga, "ILevelObject", dynamic)
    if not level_obj:
        log("Нет ILevelObject у помещения.")
        return
    level_id = int(level_obj.LevelId)
    level_mo = objects.GetById(level_id)
    ilvl = query(level_mo, Renga, "ILevel", dynamic)
    if not ilvl:
        log("Не найден уровень id=%s." % level_id)
        return

    math = app.Math
    project = app.Project

    floor_verts = get_verts(region_desc, Renga, log)
    floor_global_lines = get_lines(
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
                set_param(floor_mo, PARAM_FLOOR_THICKNESS, ft)
                pl_floor = None
                disp_fl = query(
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
                    o, ax, ay = get_axes(Renga, pl_floor)
                    loc_lines = local_lines(
                        math, Renga, floor_global_lines, o, ax, ay, log
                    )
                    loc_lines = close_lines(
                        math, Renga, loc_lines
                    )
                    comp_floor = comp_curve(math, loc_lines)
                if not comp_floor:
                    log(
                        "Контур пола в ЛСК не собран — пробую глобальный контур "
                        "(при сбое Renga проверьте контур и Placement перекрытия)."
                    )
                    comp_floor = comp_curve(math, floor_global_lines)
                if comp_floor and set_base(
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
            
            cx_c, cy_c = centroid(floor_verts)
            geom_inset = merged.get("wall_geometric_inset_from_contour", True)
            raw_inset = merged.get("wall_contour_inset_mm", None)
            if not geom_inset:
                wall_inset_mm = 0.0
            elif raw_inset is not None:
                wall_inset_mm = float(raw_inset)
            else:
                wall_inset_mm = wt * 0.5
                
            for prim in iter_outer(region_desc, Renga):
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
                    ends = curve_ends(prim)
                    if not ends:
                        continue
                    (sx, sy), (ex, ey) = ends
                    if wall_inset_mm > 1e-6:
                        nrm = inward_norm(
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
                    init_wall(wall_mo, wh, wt)
                    bl = query(
                        wall_mo, Renga, "IBaseline2DObject", dynamic
                    )
                    if bl:
                        loc = math.CreateLineSegment2D(
                            pt2d(Renga, sx, sy),
                            pt2d(Renga, ex, ey),
                        )
                        bl.SetBaseline(loc)
                elif ct == CURVE2D_ARC:
                    wall_mo = model.CreateObject(args)
                    if not wall_mo:
                        continue
                    init_wall(wall_mo, wh, wt)
                    bl = query(
                        wall_mo, Renga, "IBaseline2DObject", dynamic
                    )
                    if bl:
                        try:
                            arc_c = prim.GetCopy()
                            bl.SetBaselineInCS(
                                ident_pl(Renga), arc_c
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


def get_rooms(app, model, mode: str, explicit: List[int]) -> List[int]:
    objects = model.GetObjects()
    def is_room(mo): return mo and mo.ObjectTypeS.lower() == ENTITY_ROOM.lower()
    
    if mode == "selection":
        return [int(oid) for oid in app.Selection.GetSelectedObjects() if is_room(objects.GetById(int(oid)))]
    if mode == "all":
        try:
            return [int(mo.Id) for oid in objects.GetIds() if is_room(mo := objects.GetById(int(oid)))]
        except Exception:
            return [int(mo.Id) for i in range(objects.Count) if is_room(mo := objects.GetByIndex(i))]
    return list(explicit)


def run_batch(config_path: str, room_ids: List[int], mode: str, prefer_running: bool, log: Callable[[str], None]) -> None:
    app = connect(prefer_running)
    
    if not has_renga():
        raise RuntimeError("Не импортируется comtypes.gen.Renga. Выполните: pip install comtypes.")
    
    cfg = load_cfg(config_path)
    app.Visible = True
    if not app.HasProject:
        raise RuntimeError("В Renga нет открытого проекта.")
    
    model = app.Project.Model
    objects = model.GetObjects()
    ids = get_rooms(app, model, mode, room_ids)
    if not ids:
        raise RuntimeError("Нет помещений для обработки.")
        
    for rid in ids:
        mo = objects.GetById(int(rid))
        if mo: proc_room(app, model, objects, mo, cfg, log)
        else: log(f"Объект id={rid} не найден.")
