import asyncio
import json
import os
import sys
import httpx

project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

from Controller.bvlaichau import (
    _DEFAULT_URL,
    _DEFAULT_USER,
    _DEFAULT_PASS,
    _HEADERS,
    _extract_uuid,
    _get_default_date_range,
    get_patients_by_room,
)
from libs.login import login
from libs.export_excel import export_patients_to_excel


async def export_noi():
    log_res = await login(_DEFAULT_URL, _DEFAULT_USER, _DEFAULT_PASS)
    cookies = log_res["cookies"]
    location = log_res["location"]

    async with httpx.AsyncClient(cookies=cookies, verify=False, timeout=30.0) as client:
        uuid = await _extract_uuid(client, _DEFAULT_URL, location)
        from_date, to_date = _get_default_date_range()

        # Query phong kham noi (id 83043)
        res = await get_patients_by_room(
            client=client,
            base_url=_DEFAULT_URL,
            uuid=uuid,
            room_id="83043",
            from_date=from_date,
            to_date=to_date,
            rows=200,
        )

        patients = res.get("patients", [])
        raw_list = [p["raw"] for p in patients]

        # Export excel with both names
        export_res = export_patients_to_excel(
            raw_list,
            hospital_name="bvlaichau_phongkhamnoi",
            output_dir=project_root,
        )
        # Also save as [bvlaichau_phongkhamnoi][6].xlsx and [phongkhamnoi][6].xlsx
        print(f"Total: {len(patients)}")
        print("Export result:", export_res)

        print("\n--- PATIENT LIST ---")
        for idx, p in enumerate(patients, 1):
            r = p["raw"]
            print(
                f"{idx}. {r.get('TENBENHNHAN')} - NS: {r.get('NGAYSINH')} - GT: {r.get('GIOITINH')} "
                f"- HS: {r.get('MAHOSOBENHAN')} - DC: {r.get('DIACHI')} - SDT: {r.get('SDTBENHNHAN', '')}"
            )


if __name__ == "__main__":
    asyncio.run(export_noi())
