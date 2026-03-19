from __future__ import annotations

import time
from typing import Any
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests


class KitClient:
    def __init__(self, api_key: str, api_base: str) -> None:
        self.api_key = api_key
        self.api_base = api_base.rstrip("/")
        self.headers = {
            "X-Kit-Api-Key": self.api_key,
            "Accept": "application/json",
            "Content-Type": "application/json",
        }
        self.failed_filter_requests: list[dict[str, Any]] = []

    def _request(
        self,
        method: str,
        path: str,
        params: dict[str, Any] | None = None,
        json_body: dict[str, Any] | None = None,
        timeout: int = 30,
        max_attempts: int = 6,
    ) -> requests.Response:
        url = f"{self.api_base}{path}"
        for attempt in range(max_attempts):
            try:
                resp = requests.request(
                    method,
                    url,
                    headers=self.headers,
                    params=params,
                    json=json_body,
                    timeout=timeout,
                )
            except requests.RequestException:
                time.sleep(1 + attempt)
                continue

            if resp.status_code == 429:
                wait_s = int(resp.headers.get("Retry-After", 60))
                time.sleep(wait_s)
                continue

            if resp.status_code >= 500 and attempt < max_attempts - 1:
                time.sleep(1 + attempt)
                continue

            return resp

        raise RuntimeError(f"Kit API request failed after retries: {method} {url}")

    def _paginate_get(
        self,
        path: str,
        root_key: str,
        params: dict[str, Any] | None = None,
        per_page: int = 200,
    ) -> list[dict[str, Any]]:
        out: list[dict[str, Any]] = []
        after = None
        seen_cursors = set()

        while True:
            p = dict(params or {})
            p["per_page"] = per_page
            if after:
                p["after"] = after

            resp = self._request("GET", path, params=p)
            if resp.status_code != 200:
                raise RuntimeError(f"GET {path} failed: {resp.status_code} {resp.text[:220]}")

            payload = resp.json()
            items = payload.get(root_key, [])
            out.extend(items)

            pagination = payload.get("pagination", {})
            has_next = bool(pagination.get("has_next_page"))
            end_cursor = pagination.get("end_cursor")

            if not has_next or not end_cursor or end_cursor in seen_cursors:
                break

            seen_cursors.add(end_cursor)
            after = end_cursor

        return out

    def list_subscribers(self, status: str = "all") -> list[dict[str, Any]]:
        return self._paginate_get(
            "/subscribers",
            root_key="subscribers",
            params={"status": status},
            per_page=1000,
        )

    def list_tags(self) -> list[dict[str, Any]]:
        return self._paginate_get("/tags", root_key="tags", per_page=200)

    def list_broadcasts(self) -> list[dict[str, Any]]:
        return self._paginate_get("/broadcasts", root_key="broadcasts", per_page=200)

    def list_tag_subscribers(self, tag_id: int) -> set[str]:
        subs = self._paginate_get(f"/tags/{tag_id}/subscribers", root_key="subscribers", per_page=1000)
        emails = {
            str(s.get("email_address", "")).strip().lower()
            for s in subs
            if s.get("email_address")
        }
        return emails

    def get_broadcast_stats(self, broadcast_id: int) -> dict[str, Any]:
        resp = self._request("GET", f"/broadcasts/{broadcast_id}/stats")
        if resp.status_code != 200:
            return {}
        data = resp.json().get("broadcast", {})
        return data.get("stats", {}) or {}

    def _filter_with_payload(self, payload_base: dict[str, Any]) -> tuple[set[str], int, str]:
        results: set[str] = set()
        after = None
        seen_cursors = set()

        while True:
            payload = dict(payload_base)
            if after:
                payload["after"] = after

            resp = self._request("POST", "/subscribers/filter", json_body=payload)
            if resp.status_code != 200:
                return set(), resp.status_code, resp.text[:220]

            payload_json = resp.json()
            subscribers = payload_json.get("subscribers", [])
            for s in subscribers:
                e = s.get("email_address")
                if e:
                    results.add(str(e).strip().lower())

            pagination = payload_json.get("pagination", {})
            has_next = bool(pagination.get("has_next_page"))
            end_cursor = pagination.get("end_cursor")
            if not has_next or not end_cursor or end_cursor in seen_cursors:
                break

            seen_cursors.add(end_cursor)
            after = end_cursor

        return results, 200, ""

    def _filter_chunk(self, event_type: str, ids_chunk: list[int]) -> set[str]:
        payload_base = {
            "all": [
                {
                    "type": event_type,
                    "count_greater_than": 0,
                    "any": [{"type": "broadcasts", "ids": [int(x) for x in ids_chunk]}],
                }
            ]
        }

        results, status_code, err_preview = self._filter_with_payload(payload_base)
        if status_code == 200:
            return results

        if len(ids_chunk) > 1:
            mid = len(ids_chunk) // 2
            return self._filter_chunk(event_type, ids_chunk[:mid]) | self._filter_chunk(
                event_type, ids_chunk[mid:]
            )

        self.failed_filter_requests.append(
            {
                "event_type": event_type,
                "broadcast_id": ids_chunk[0],
                "status_code": status_code,
                "error_preview": err_preview,
            }
        )
        return set()

    def filter_subscribers_by_broadcast_event(
        self,
        event_type: str,
        broadcast_ids: list[int],
        chunk_size: int = 12,
    ) -> set[str]:
        if not broadcast_ids:
            return set()

        dedup_ids = sorted({int(x) for x in broadcast_ids})
        all_emails: set[str] = set()
        chunks = [dedup_ids[i : i + chunk_size] for i in range(0, len(dedup_ids), chunk_size)]
        if not chunks:
            return set()

        with ThreadPoolExecutor(max_workers=min(8, len(chunks))) as executor:
            future_map = {
                executor.submit(self._filter_chunk, event_type, chunk): chunk
                for chunk in chunks
            }
            for future in as_completed(future_map):
                try:
                    all_emails |= future.result()
                except Exception:
                    chunk = future_map[future]
                    if len(chunk) == 1:
                        self.failed_filter_requests.append(
                            {
                                "event_type": event_type,
                                "broadcast_id": chunk[0],
                                "status_code": "client_error",
                                "error_preview": "Unhandled exception during filter chunk execution.",
                            }
                        )

        return all_emails

    def filter_subscribers_by_event_date(
        self,
        event_type: str,
        after_iso_date: str | None = None,
        before_iso_date: str | None = None,
    ) -> set[str]:
        condition: dict[str, Any] = {"type": event_type, "count_greater_than": 0}
        if after_iso_date:
            condition["after"] = after_iso_date
        if before_iso_date:
            condition["before"] = before_iso_date

        payload_base = {"all": [condition]}
        results, status_code, err_preview = self._filter_with_payload(payload_base)
        if status_code != 200:
            self.failed_filter_requests.append(
                {
                    "event_type": event_type,
                    "broadcast_id": None,
                    "status_code": status_code,
                    "error_preview": err_preview,
                }
            )
            return set()
        return results

    def get_subscriber_tags(self, subscriber_id: int) -> list[str]:
        tags = self._paginate_get(f"/subscribers/{subscriber_id}/tags", root_key="tags", per_page=1000)
        return [str(t.get("name", "")).strip() for t in tags if t.get("name")]
