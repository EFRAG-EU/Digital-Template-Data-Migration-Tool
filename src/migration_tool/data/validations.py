import json
from dataclasses import dataclass, field
from typing import Required, TypedDict


def convertToNone(string: str) -> str | None:
    NONESTRING = "None"
    if string == NONESTRING:
        return None
    return string


class AnchorJson(TypedDict, total=False):
    label: Required[str]
    cols: Required[str]
    rows: int
    rowOffset: int


class RefsJson(TypedDict, total=False):
    sheet: Required[str]
    range: str | None
    anchor: AnchorJson


@dataclass
class Anchor:
    """Locates cells by the row holding a `=template_label_…` cell, so they move
    with that row when a template release inserts rows above it."""

    label: str  # the template_label_… name, or plain cell text
    cols: str  # "D" or "D:H"
    rows: int = 1
    rowOffset: int = 0  # rows below the label's row

    def toJson(self) -> AnchorJson:
        json: AnchorJson = {"label": self.label, "cols": self.cols}
        if self.rows != 1:
            json["rows"] = self.rows
        if self.rowOffset:
            json["rowOffset"] = self.rowOffset
        return json


@dataclass
class Refs:
    label: str
    sheet: str
    range: str | None  # for an anchor: None until resolved
    anchor: Anchor | None = None

    @classmethod
    def fromJson(cls, version: str, label: str, ref: RefsJson) -> "Refs":
        where = f"missing_NR {version} '{label}'"
        match sorted(ref):
            case ["range", "sheet"]:
                return cls(label, ref["sheet"], ref["range"])
            case ["anchor", "sheet"]:
                anchor = ref["anchor"]
                if (
                    not {"label", "cols"}
                    <= set(anchor)
                    <= set(AnchorJson.__annotations__)
                ):
                    raise ValueError(
                        f"{where}: anchor needs 'label' and 'cols' (and optionally 'rows', 'rowOffset'), got {sorted(anchor)}"
                    )
                return cls(label, ref["sheet"], None, Anchor(**anchor))
        raise ValueError(
            f"{where}: expected 'sheet' and either 'range' or 'anchor', got {sorted(ref)}"
        )

    def toJson(self) -> RefsJson:
        if self.anchor:
            return {"sheet": self.sheet, "anchor": self.anchor.toJson()}
        return {"sheet": self.sheet, "range": self.range}


@dataclass
class Version:
    version: str
    missing_data: list[Refs]


@dataclass
class VersionCollection:
    versions: list[Version] = field(default_factory=list)

    def add_version(self, version: str, labels_refs: dict[str, RefsJson]) -> Version:
        trios = [
            Refs.fromJson(version, label, ref) for label, ref in labels_refs.items()
        ]
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

    def getRefs(self, version: str) -> list[Refs]:
        return [
            ref for v in self.versions if v.version == version for ref in v.missing_data
        ]

    def get_ranges(self, version: str) -> list[str | None]:
        if unresolved := [
            ref.label for ref in self.getRefs(version) if ref.anchor and not ref.range
        ]:
            raise ValueError(f"missing_NR {version}: unresolved anchors {unresolved}")
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
            v.version: {ref.label: ref.toJson() for ref in v.missing_data}
            for v in self.versions
        }

    def write_json(self, filepath: str) -> None:
        with open(filepath, "w") as json_file:
            json.dump(self.dump_simple(), json_file, indent=4)
