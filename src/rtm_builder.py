"""
rtm_builder.py — построение матрицы трассировки требований.

Строит связи:
  BR → FR → TC

и формирует полный RTM-граф со всеми перекрёстными ссылками.
"""


def build_rtm(data: dict) -> dict:
    """
    Принимает dict с ключами 'br', 'fr', 'tc'.
    Возвращает enriched RTM dict.
    """
    br_map = {r["id"]: r for r in data["br"]}
    fr_map = {r["id"]: r for r in data["fr"]}
    tc_map = {r["id"]: r for r in data["tc"]}

    # FR → список TC, покрывающих данный FR
    fr_to_tcs: dict[str, list[str]] = {fr_id: [] for fr_id in fr_map}
    for tc in data["tc"]:
        for fr_id in tc["fr_refs"]:
            if fr_id in fr_to_tcs:
                fr_to_tcs[fr_id].append(tc["id"])

    # BR → список FR, реализующих данный BR
    br_to_frs: dict[str, list[str]] = {br_id: [] for br_id in br_map}
    for fr in data["fr"]:
        for br_id in fr["br_refs"]:
            if br_id in br_to_frs:
                br_to_frs[br_id].append(fr["id"])

    # BR → список TC (транзитивно через FR)
    br_to_tcs: dict[str, list[str]] = {}
    for br_id in br_map:
        tcs: list[str] = []
        for fr_id in br_to_frs[br_id]:
            tcs.extend(fr_to_tcs.get(fr_id, []))
        br_to_tcs[br_id] = list(dict.fromkeys(tcs))  # deduplicate, keep order

    # Orphan detection
    orphan_fr = [fr["id"] for fr in data["fr"] if not fr["br_refs"]]
    orphan_tc = [tc["id"] for tc in data["tc"] if not tc["fr_refs"]]

    # Build flat RTM rows (one row per BR–FR–TC combination)
    rtm_rows = _build_rows(data["br"], br_to_frs, fr_map, fr_to_tcs, tc_map)

    return {
        "br_map":    br_map,
        "fr_map":    fr_map,
        "tc_map":    tc_map,
        "br_to_frs": br_to_frs,
        "fr_to_tcs": fr_to_tcs,
        "br_to_tcs": br_to_tcs,
        "orphan_fr": orphan_fr,
        "orphan_tc": orphan_tc,
        "rtm_rows":  rtm_rows,
    }


def _build_rows(
    br_list,
    br_to_frs: dict,
    fr_map: dict,
    fr_to_tcs: dict,
    tc_map: dict,
) -> list[dict]:
    """
    Разворачивает иерархию BR→FR→TC в плоский список строк.
    Каждая строка содержит данные одного сочетания BR/FR/TC.
    """
    rows = []
    for br in br_list:
        br_id   = br["id"]
        fr_ids  = br_to_frs.get(br_id, [])

        if not fr_ids:
            # BR без FR
            rows.append(_row(br, None, None))
            continue

        for fr_id in fr_ids:
            fr = fr_map.get(fr_id)
            if fr is None:
                rows.append(_row(br, {"id": fr_id, "title": "⚠ Не найден"}, None))
                continue

            tc_ids = fr_to_tcs.get(fr_id, [])
            if not tc_ids:
                # FR без TC
                rows.append(_row(br, fr, None))
                continue

            for tc_id in tc_ids:
                tc = tc_map.get(tc_id)
                if tc is None:
                    rows.append(_row(br, fr, {"id": tc_id, "title": "⚠ Не найден", "result": ""}))
                else:
                    rows.append(_row(br, fr, tc))

    return rows


def _row(br: dict, fr: dict | None, tc: dict | None) -> dict:
    return {
        "br_id":       br["id"],
        "br_title":    br["title"],
        "br_priority": br.get("priority", ""),
        "br_category": br.get("category", ""),
        "br_status":   br.get("status", "Active"),
        "fr_id":       fr["id"]    if fr else "",
        "fr_title":    fr["title"] if fr else "",
        "fr_type":     fr.get("type", "")      if fr else "",
        "fr_component":fr.get("component", "") if fr else "",
        "fr_status":   fr.get("status", "")    if fr else "",
        "tc_id":       tc["id"]     if tc else "",
        "tc_title":    tc["title"]  if tc else "",
        "tc_type":     tc.get("type", "")     if tc else "",
        "tc_result":   tc.get("result", "")   if tc else "",
        "tc_priority": tc.get("priority", "") if tc else "",
    }
