from dataclasses import dataclass


@dataclass(slots=True)
class RepairStats:
    checked: int = 0
    fixed: int = 0
    missing: int = 0
    ambiguous: int = 0


@dataclass(slots=True)
class RepairResult:
    stats: RepairStats
    missing_files: list[str]
    log_lines: list[str]