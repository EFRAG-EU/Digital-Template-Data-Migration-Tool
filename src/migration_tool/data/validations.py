import json
from dataclasses import dataclass, field


def convertToNone(string: str) -> str | None:
    NONESTRING = "None"
    if string == NONESTRING:
        return None
    return string


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

    def add_version(self, version: str, labels_refs: dict[str, str]) -> Version:
        labels = [lab for lab in labels_refs.keys()]
        sheets = [ref.split("!")[0] for ref in labels_refs.values()]
        ranges = [convertToNone(ref.split("!")[1]) for ref in labels_refs.values()]
        trios = []
        for i in range(len(labels)):
            refsToAppend = Refs(
                label=labels[i],
                sheet=sheets[i],
                range=ranges[i],
            )
            trios.append(refsToAppend)
        v = Version(version=version, missing_data=trios)
        self.versions.append(v)
        return v

    def get_version(self, version: str) -> dict[str, str]:
        return {
            trios.label: f"{trios.sheet}!{trios.range}"
            for v in self.versions
            for trios in v.missing_data
            if v.version == version
        }

    def get_labels(self, version: str) -> list[str]:
        labelRefs = self.get_version(version)
        return [label for label in labelRefs.keys()]

    def get_sheets(self, version: str) -> list[str]:
        labelRefs = self.get_version(version)
        return [ref.split("!")[0] for ref in labelRefs.values()]

    def get_ranges(self, version: str) -> list[str | None]:
        labelRefs = self.get_version(version)
        return [convertToNone(ref.split("!")[1]) for ref in labelRefs.values()]

    @classmethod
    def from_dict(cls, data: dict[str, dict[str, str]]) -> "VersionCollection":
        """
        Build from: {"1.0.0": {"stable": "refs/tags/v1.0.0"}, ...}
        """
        collection = cls()
        for version, labels_refs in data.items():
            collection.add_version(version, labels_refs)
        return collection

    def dump_simple(self) -> dict[str, dict[str, str]]:
        """
        Dump as: [{"1.0.0": {"stable": "refs/tags/v1.0.0", ...}}, ...]
        """
        return {v.version: self.get_version(v.version) for v in self.versions}

    def write_json(self, filepath: str) -> None:
        with open(filepath, "w") as json_file:
            json.dump(self.dump_simple(), json_file, indent=4)
