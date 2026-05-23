from __future__ import annotations

import json
import logging
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Optional


CHECKPOINT_VERSION = 1


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


@dataclass
class CheckpointManager:
    path: Path
    logger: logging.Logger

    def __post_init__(self) -> None:
        self.path = Path(self.path)
        self.path.parent.mkdir(parents=True, exist_ok=True)

    def load(self) -> Optional[Dict[str, Any]]:
        if not self.path.exists():
            return None

        try:
            payload = json.loads(self.path.read_text(encoding="utf-8"))
        except Exception as exc:
            self.logger.error("Unable to load checkpoint %s: %s", self.path, exc, exc_info=True)
            return None

        if not isinstance(payload, dict):
            self.logger.warning("Checkpoint %s is not a valid JSON object.", self.path)
            return None

        return payload

    def save(
        self,
        *,
        data: Dict[str, Any],
        processed_ids: set[str],
        last_species: Optional[Dict[str, str]],
        stats: Dict[str, Any],
        output_path: str,
        partial_output_path: str,
        completed: bool = False,
    ) -> None:
        payload = {
            "version": CHECKPOINT_VERSION,
            "updated_at": utc_now_iso(),
            "completed": completed,
            "output_path": output_path,
            "partial_output_path": partial_output_path,
            "stats": stats,
            "last_species": last_species or {},
            "processed_avibase_ids": sorted(processed_ids),
            "data": data,
        }

        tmp_path = self.path.with_suffix(".tmp")
        tmp_path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        tmp_path.replace(self.path)

        self.logger.info(
            "Checkpoint saved to %s | processed=%s | completed=%s",
            self.path,
            len(processed_ids),
            completed,
        )
