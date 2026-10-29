import json
from dataclasses import dataclass, field
from typing import TypedDict


def convertToNone(string: str) -> str | None:
    NONESTRING = "None"
    if string == NONESTRING:
        return None
    return string


class RefsJson(TypedDict):
    sheet: str
    range: str | None


@dataclass
class Refs:
    label: str
    sheet: str
    range: str | None


@dataclass
class Version:
    version: str
    missing_data: list[Refs]


@dataclass
class VersionCollection:
    versions: list[Version] = field(default_factory=list)

    def add_version(self, version: str, labels_refs: dict[str, RefsJson]) -> Version:
        trios = []
        for label, ref in labels_refs.items():
            if set(ref) != {"sheet", "range"}:
                raise ValueError(
                    f"missing_NR {version} '{label}': expected keys 'sheet' and 'range', got {sorted(ref)}"
                )
            trios.append(Refs(label=label, sheet=ref["sheet"], range=ref["range"]))
        v = Version(version=version, missing_data=trios)
        self.versions.append(v)
        return v

    def get_versions(self) -> list[str]:
        return [v.version for v in self.versions]

    def get_version(self, version: str) -> dict[str, str]:
        return {
            trios.label: f"{trios.sheet}!{trios.range}"
            for v in self.versions
            for trios in v.missing_data
            if v.version == version
        }

    def get_labels(self, version: str) -> list[str]:
        labelRefs = self.get_version(version)
        return list(labelRefs)

    def get_sheets(self, version: str) -> list[str]:
        labelRefs = self.get_version(version)
        return [ref.split("!")[0] for ref in labelRefs.values()]

    def get_ranges(self, version: str) -> list[str | None]:
        labelRefs = self.get_version(version)
        return [convertToNone(ref.split("!")[1]) for ref in labelRefs.values()]

    def get_refs_from_label(self, label: str) -> dict[str, str]:
        return {
            version.version: f"{ref.sheet}!{ref.range}"
            for version in self.versions
            for ref in version.missing_data
            if ref.label == label
        }

    def get_refs_from_sheet(self, sheet: str) -> dict[str, dict[str, str]]:
        result: dict[str, dict[str, str]] = {}

        for version in self.versions:
            for ref in version.missing_data:
                if ref.sheet == sheet:
                    result.setdefault(version.version, {})
                    result[version.version][ref.label] = f"{ref.sheet}!{ref.range}"

        return result

    @classmethod
    def from_dict(cls, data: dict[str, dict[str, RefsJson]]) -> "VersionCollection":
        """
        Build from: {"1.0.0": {"StartDate": {"sheet": "General Information", "range": "$D$6:$E$8"}}, ...}
        """
        collection = cls()
        for version, labels_refs in data.items():
            collection.add_version(version, labels_refs)
        return collection

    def dump_simple(self) -> dict[str, dict[str, RefsJson]]:
        """
        Dump in the same shape from_dict() reads.
        """
        return {
            v.version: {
                ref.label: {"sheet": ref.sheet, "range": ref.range}
                for ref in v.missing_data
            }
            for v in self.versions
        }

    def write_json(self, filepath: str) -> None:
        with open(filepath, "w") as json_file:
            json.dump(self.dump_simple(), json_file, indent=4)
