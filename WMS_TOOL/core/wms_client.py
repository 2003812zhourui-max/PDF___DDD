from __future__ import annotations

from pathlib import Path
from typing import Any


LIST_API = "/gateway/wms/blDelivery/page"


def build_list_payload(
    *,
    current: int,
    size: int,
    start_time: str,
    end_time: str,
    wh_code: str,
    status: str,
    channel: str = "",
) -> dict[str, Any]:
    return {
        "appendixFlag": "",
        "areaCodes": [],
        "categoryIdList": [],
        "cellNos": [],
        "codeType": "barcode",
        "countKind": "orderWeight",
        "countryRegionCodes": "",
        "current": current,
        "customerCodes": "",
        "endTime": end_time,
        "expressFlag": "",
        "expressPrintStatus": "",
        "forecastStatus": "",
        "logisticsCarrier": "",
        "logisticsChannel": channel,
        "orderCount": "",
        "orderNoType": "sourceNo",
        "orderSourceList": [],
        "productPackType": "",
        "receiver": "",
        "relatedReturnOrder": "",
        "salesPlatform": "",
        "size": size,
        "skuQtyStrList": [],
        "sourceNoLists": [],
        "startTime": start_time,
        "status": status,
        "timeType": "createTime",
        "unitMark": 0,
        "varietyType": "",
        "weightCountEnd": "",
        "weightCountStart": "",
        "whCode": wh_code,
        "withVas": "",
    }


def query_totals(
    *,
    start_time: str,
    end_time: str,
    wh_codes: list[str],
    statuses: list[str],
    storage_state: str,
    channel: str = "",
) -> dict[str, Any]:
    from batch_download_wms_pdfs import fetch_json_http_with_auth_retry, session_from_storage_state

    totals: dict[str, Any] = {"by_warehouse": {}, "total": 0}
    state_path = str(Path(storage_state).expanduser().resolve())
    for wh_code in wh_codes:
        session, auth_values = session_from_storage_state(state_path, wh_code, "auto")
        session.headers["whcode"] = wh_code
        wh_total = 0
        status_totals: dict[str, int] = {}
        for status in statuses:
            payload = build_list_payload(
                current=1,
                size=1,
                start_time=start_time,
                end_time=end_time,
                wh_code=wh_code,
                status=status,
                channel=channel,
            )
            data = fetch_json_http_with_auth_retry(session, auth_values, LIST_API, method="POST", payload=payload)
            body = data.get("data", {}) if isinstance(data, dict) else {}
            count = int(body.get("total") or 0) if isinstance(body, dict) else 0
            status_totals[status] = count
            wh_total += count
        totals["by_warehouse"][wh_code] = {"total": wh_total, "statuses": status_totals}
        totals["total"] += wh_total
    return totals
