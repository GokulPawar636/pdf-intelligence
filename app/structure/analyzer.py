from __future__ import annotations

import json
import math
from dataclasses import dataclass
from pathlib import Path
from statistics import median
from typing import Any

from app.models.structure import (
    PageStructure,
    Region,
    StructuredDocument,
    TableCandidate,
    TableCell,
)
from app.models.document import RawDocument


@dataclass(frozen=True)
class Word:
    """
    Internal representation of one PDF word.
    """

    index: int
    text: str

    x0: float
    y0: float
    x1: float
    y1: float

    block: int
    line: int
    word: int

    @property
    def width(self) -> float:
        return max(
            0.0,
            self.x1 - self.x0,
        )

    @property
    def height(self) -> float:
        return max(
            0.0,
            self.y1 - self.y0,
        )

    @property
    def cx(self) -> float:
        return (
            self.x0 + self.x1
        ) / 2.0

    @property
    def cy(self) -> float:
        return (
            self.y0 + self.y1
        ) / 2.0


@dataclass
class Line:
    """
    Visually reconstructed line of text.
    """

    words: list[Word]

    @property
    def x0(self) -> float:
        return min(
            w.x0 for w in self.words
        )

    @property
    def y0(self) -> float:
        return min(
            w.y0 for w in self.words
        )

    @property
    def x1(self) -> float:
        return max(
            w.x1 for w in self.words
        )

    @property
    def y1(self) -> float:
        return max(
            w.y1 for w in self.words
        )

    @property
    def cy(self) -> float:
        return (
            self.y0 + self.y1
        ) / 2.0

    @property
    def height(self) -> float:
        return max(
            0.0,
            self.y1 - self.y0,
        )

    @property
    def text(self) -> str:
        return " ".join(
            word.text
            for word in sorted(
                self.words,
                key=lambda x: x.x0,
            )
        )


class StructureAnalyzer:
    """
    Dynamic document structure analyzer.

    It DOES NOT look for:

        page == 2
        "Feature Table"
        "Nom"
        "Meas"
        "Tol"
        "Ballooning"

    Instead it uses:

        word coordinates
        line alignment
        vertical spacing
        repeated x positions
        row density
        column density
    """

    def __init__(
        self,
        *,
        region_gap_factor: float = 1.8,
        max_region_gap_points: float = 48.0,
        row_tolerance_factor: float = 0.55,
        column_tolerance_factor: float = 1.5,
        min_table_rows: int = 2,
        min_table_columns: int = 2,
        min_table_density: float = 0.45,
        min_table_words_per_row: int = 2,
        max_table_row_gap_factor: float = 2.8,
        anchor_match_tolerance_factor: float = 1.8,
    ) -> None:

        self.region_gap_factor = (
            region_gap_factor
        )

        self.max_region_gap_points = (
            max_region_gap_points
        )

        self.row_tolerance_factor = (
            row_tolerance_factor
        )

        self.column_tolerance_factor = (
            column_tolerance_factor
        )

        self.min_table_rows = (
            min_table_rows
        )

        self.min_table_columns = (
            min_table_columns
        )

        self.min_table_density = (
            min_table_density
        )

        self.min_table_words_per_row = (
            min_table_words_per_row
        )

        self.max_table_row_gap_factor = (
            max_table_row_gap_factor
        )

        self.anchor_match_tolerance_factor = (
            anchor_match_tolerance_factor
        )

    # ---------------------------------------------------------
    # Geometry
    # ---------------------------------------------------------

    @staticmethod
    def _bbox(
        words: list[Word],
    ) -> dict[str, float]:

        return {
            "x0": round(
                min(w.x0 for w in words),
                4,
            ),
            "y0": round(
                min(w.y0 for w in words),
                4,
            ),
            "x1": round(
                max(w.x1 for w in words),
                4,
            ),
            "y1": round(
                max(w.y1 for w in words),
                4,
            ),
        }

    @staticmethod
    def _bbox_area(
        bbox: dict[str, float],
    ) -> float:

        return (
            max(
                0.0,
                bbox["x1"] - bbox["x0"],
            )
            *
            max(
                0.0,
                bbox["y1"] - bbox["y0"],
            )
        )

    @staticmethod
    def _intersection_area(
        a: dict[str, float],
        b: dict[str, float],
    ) -> float:

        x0 = max(
            a["x0"],
            b["x0"],
        )

        y0 = max(
            a["y0"],
            b["y0"],
        )

        x1 = min(
            a["x1"],
            b["x1"],
        )

        y1 = min(
            a["y1"],
            b["y1"],
        )

        if x1 <= x0 or y1 <= y0:
            return 0.0

        return (
            (x1 - x0)
            *
            (y1 - y0)
        )

    # ---------------------------------------------------------
    # Load words
    # ---------------------------------------------------------

    @staticmethod
    def _load_words(
        page: dict[str, Any],
    ) -> list[Word]:

        result: list[Word] = []

        for element in page.get(
            "elements",
            [],
        ):

            if element.get("type") != "word":
                continue

            bbox = element.get(
                "bbox"
            ) or {}

            text = str(
                element.get(
                    "text",
                    "",
                )
            ).strip()

            if not text:
                continue

            try:

                result.append(
                    Word(
                        index=len(result),

                        text=text,

                        x0=float(
                            bbox["x0"]
                        ),

                        y0=float(
                            bbox["y0"]
                        ),

                        x1=float(
                            bbox["x1"]
                        ),

                        y1=float(
                            bbox["y1"]
                        ),

                        block=int(
                            element.get(
                                "block_number",
                                -1,
                            )
                        ),

                        line=int(
                            element.get(
                                "line_number",
                                -1,
                            )
                        ),

                        word=int(
                            element.get(
                                "word_number",
                                -1,
                            )
                        ),
                    )
                )

            except (
                KeyError,
                TypeError,
                ValueError,
            ):
                continue

        return result

    # ---------------------------------------------------------
    # Word → lines
    # ---------------------------------------------------------

    def _cluster_lines(
        self,
        words: list[Word],
    ) -> list[Line]:

        if not words:
            return []

        ordered = sorted(
            words,
            key=lambda w: (
                w.cy,
                w.x0,
            ),
        )

        heights = [
            w.height
            for w in words
            if w.height > 0
        ]

        typical_height = (
            median(heights)
            if heights
            else 8.0
        )

        tolerance = max(
            1.5,
            typical_height
            * self.row_tolerance_factor,
        )

        clusters: list[
            list[Word]
        ] = []

        centers: list[float] = []

        for word in ordered:

            best_index = None

            best_distance = (
                float("inf")
            )

            for i, center in enumerate(
                centers
            ):

                distance = abs(
                    word.cy - center
                )

                if (
                    distance <= tolerance
                    and
                    distance < best_distance
                ):

                    best_index = i
                    best_distance = distance

            if best_index is None:

                clusters.append(
                    [word]
                )

                centers.append(
                    word.cy
                )

            else:

                clusters[
                    best_index
                ].append(word)

                centers[
                    best_index
                ] = (
                    sum(
                        w.cy
                        for w in clusters[
                            best_index
                        ]
                    )
                    /
                    len(
                        clusters[
                            best_index
                        ]
                    )
                )

        lines = []

        for group in sorted(
            clusters,
            key=lambda g: min(
                w.cy for w in g
            ),
        ):

            lines.append(
                Line(
                    words=sorted(
                        group,
                        key=lambda w: w.x0,
                    )
                )
            )

        return lines

    # ---------------------------------------------------------
    # Text regions
    # ---------------------------------------------------------

    def _group_text_regions(
        self,
        lines: list[Line],
    ) -> list[list[Line]]:

        if not lines:
            return []

        heights = [
            line.height
            for line in lines
            if line.height > 0
        ]

        typical_height = (
            median(heights)
            if heights
            else 8.0
        )

        allowed_gap = min(
            self.max_region_gap_points,
            max(
                8.0,
                typical_height
                * self.region_gap_factor,
            ),
        )

        regions = [
            [lines[0]]
        ]

        for line in lines[1:]:

            previous = regions[-1][-1]

            vertical_gap = max(
                0.0,
                line.y0 - previous.y1,
            )

            horizontal_gap = max(
                0.0,
                max(
                    line.x0,
                    previous.x0,
                )
                -
                min(
                    line.x1,
                    previous.x1,
                ),
            )

            horizontal_overlap = (
                min(
                    line.x1,
                    previous.x1,
                )
                -
                max(
                    line.x0,
                    previous.x0,
                )
            )

            if (
                vertical_gap <= allowed_gap
                and
                (
                    horizontal_overlap >= 0
                    or
                    horizontal_gap
                    <= allowed_gap * 2
                )
            ):

                regions[-1].append(
                    line
                )

            else:

                regions.append(
                    [line]
                )

        return regions

    def _make_regions(
        self,
        page_number: int,
        lines: list[Line],
    ) -> list[Region]:

        regions = []

        grouped = (
            self._group_text_regions(
                lines
            )
        )

        for index, group in enumerate(
            grouped,
            start=1,
        ):

            words = [
                word
                for line in group
                for word in line.words
            ]

            if not words:
                continue

            regions.append(
                Region(
                    region_id=(
                        f"p{page_number:03d}"
                        f"_r{index:03d}"
                    ),

                    page_number=page_number,

                    region_type="text_region",

                    bbox=self._bbox(
                        words
                    ),

                    element_indexes=[
                        word.index
                        for word in words
                    ],

                    text="\n".join(
                        line.text
                        for line in group
                    ),

                    confidence=0.90,

                    evidence={
                        "method":
                            "spatial_text_clustering",

                        "line_count":
                            len(group),

                        "word_count":
                            len(words),
                    },
                )
            )

        return regions

    # ---------------------------------------------------------
    # Table detection
    # ---------------------------------------------------------

    def _line_anchor_similarity(
        self,
        a: Line,
        b: Line,
        tolerance: float,
    ) -> float:

        if (
            len(a.words) < 2
            or len(b.words) < 2
        ):
            return 0.0

        matched = 0

        used: set[int] = set()

        for word_a in a.words:

            best = None

            best_distance = (
                float("inf")
            )

            for j, word_b in enumerate(
                b.words
            ):

                if j in used:
                    continue

                distance = abs(
                    word_a.x0
                    -
                    word_b.x0
                )

                if (
                    distance <= tolerance
                    and
                    distance < best_distance
                ):

                    best = j
                    best_distance = distance

            if best is not None:

                used.add(best)
                matched += 1

        return (
            matched
            /
            max(
                len(a.words),
                len(b.words),
            )
        )

    def _is_table_like_line(
        self,
        line: Line,
    ) -> bool:

        return (
            len(line.words)
            >=
            self.min_table_words_per_row
        )

    def _find_table_row_runs(
        self,
        lines: list[Line],
    ) -> list[list[Line]]:

        if len(lines) < self.min_table_rows:
            return []

        heights = [
            word.height
            for line in lines
            for word in line.words
            if word.height > 0
        ]

        typical_height = (
            median(heights)
            if heights
            else 8.0
        )

        max_gap = max(
            12.0,
            typical_height
            * self.max_table_row_gap_factor,
        )

        anchor_tolerance = max(
            3.0,
            typical_height
            * self.anchor_match_tolerance_factor,
        )

        runs = []

        current: list[Line] = []

        for line in lines:

            if not self._is_table_like_line(
                line
            ):

                if (
                    len(current)
                    >= self.min_table_rows
                ):

                    runs.append(
                        current
                    )

                current = []

                continue

            if not current:

                current = [
                    line
                ]

                continue

            previous = current[-1]

            vertical_gap = max(
                0.0,
                line.y0
                -
                previous.y1,
            )

            similarity = (
                self._line_anchor_similarity(
                    previous,
                    line,
                    anchor_tolerance,
                )
            )

            if (
                vertical_gap <= max_gap
                and
                similarity >= 0.45
            ):

                current.append(
                    line
                )

            else:

                if (
                    len(current)
                    >= self.min_table_rows
                ):

                    runs.append(
                        current
                    )

                current = [
                    line
                ]

        if (
            len(current)
            >= self.min_table_rows
        ):

            runs.append(
                current
            )

        return runs

    # ---------------------------------------------------------
    # Column detection
    # ---------------------------------------------------------

    def _discover_columns(
        self,
        lines: list[Line],
    ) -> list[float]:

        words = [
            word
            for line in lines
            for word in line.words
        ]

        if not words:
            return []

        heights = [
            word.height
            for word in words
            if word.height > 0
        ]

        typical_height = (
            median(heights)
            if heights
            else 8.0
        )

        tolerance = max(
            3.0,
            typical_height
            * self.column_tolerance_factor,
        )

        anchors: list[float] = []

        counts: list[int] = []

        for word in sorted(
            words,
            key=lambda w: w.x0,
        ):

            best = None

            best_distance = (
                float("inf")
            )

            for i, anchor in enumerate(
                anchors
            ):

                distance = abs(
                    word.x0
                    -
                    anchor
                )

                if (
                    distance <= tolerance
                    and
                    distance < best_distance
                ):

                    best = i
                    best_distance = distance

            if best is None:

                anchors.append(
                    word.x0
                )

                counts.append(
                    1
                )

            else:

                counts[best] += 1

                anchors[best] = (
                    (
                        anchors[best]
                        *
                        (counts[best] - 1)
                    )
                    +
                    word.x0
                ) / counts[best]

        selected = []

        minimum_hits = max(
            2,
            math.ceil(
                len(lines)
                * 0.45
            ),
        )

        for anchor in anchors:

            row_hits = sum(
                1
                for line in lines
                if any(
                    abs(
                        word.x0
                        -
                        anchor
                    )
                    <= tolerance
                    for word in line.words
                )
            )

            if row_hits >= minimum_hits:
                selected.append(
                    anchor
                )

        merged = []

        for anchor in sorted(
            selected
        ):

            if (
                not merged
                or
                abs(
                    anchor
                    -
                    merged[-1]
                )
                >
                tolerance
            ):

                merged.append(
                    anchor
                )

            else:

                merged[-1] = (
                    merged[-1]
                    +
                    anchor
                ) / 2.0

        return merged

    # ---------------------------------------------------------
    # Build table
    # ---------------------------------------------------------

    def _candidate_from_run(
        self,
        page_number: int,
        lines: list[Line],
        candidate_index: int,
    ) -> TableCandidate | None:

        columns = (
            self._discover_columns(
                lines
            )
        )

        if (
            len(columns)
            <
            self.min_table_columns
        ):
            return None

        heights = [
            word.height
            for line in lines
            for word in line.words
            if word.height > 0
        ]

        typical_height = (
            median(heights)
            if heights
            else 8.0
        )

        if len(columns) == 2:
            separated_rows = sum(
                any(right.x0 - left.x1 > typical_height * 1.5
                    for left, right in zip(sorted(line.words, key=lambda w: w.x0),
                                           sorted(line.words, key=lambda w: w.x0)[1:]))
                for line in lines
            )
            if separated_rows < self.min_table_rows:
                return None

        assignment_tolerance = max(
            4.0,
            typical_height
            *
            self.column_tolerance_factor
            *
            1.5,
        )

        occupied: set[
            tuple[int, int]
        ] = set()

        cells: list[
            TableCell
        ] = []

        for row_index, line in enumerate(
            lines
        ):

            by_column: dict[
                int,
                list[Word]
            ] = {}

            for word in line.words:

                distances = [
                    abs(
                        word.x0
                        -
                        column_x
                    )
                    for column_x in columns
                ]

                column_index = min(
                    range(
                        len(columns)
                    ),
                    key=lambda i:
                        distances[i],
                )

                if (
                    distances[
                        column_index
                    ]
                    <= assignment_tolerance
                ):

                    by_column.setdefault(
                        column_index,
                        [],
                    ).append(
                        word
                    )

                    occupied.add(
                        (
                            row_index,
                            column_index,
                        )
                    )

            for (
                column_index,
                cell_words,
            ) in sorted(
                by_column.items()
            ):

                cell_words = sorted(
                    cell_words,
                    key=lambda w:
                        w.x0,
                )

                cells.append(
                    TableCell(
                        row=row_index,

                        column=column_index,

                        text=" ".join(
                            word.text
                            for word in cell_words
                        ),

                        bbox=self._bbox(
                            cell_words
                        ),

                        source_word_indexes=[
                            word.index
                            for word in cell_words
                        ],
                    )
                )

        row_count = len(lines)

        column_count = len(columns)

        density = (
            len(occupied)
            /
            float(
                row_count
                *
                column_count
            )
        )

        if (
            density
            <
            self.min_table_density
        ):
            return None

        populated_rows = sum(
            1
            for row in range(
                row_count
            )
            if sum(
                1
                for r, _ in occupied
                if r == row
            )
            >= self.min_table_columns
        )

        if (
            populated_rows
            <
            self.min_table_rows
        ):
            return None

        all_words = [
            word
            for line in lines
            for word in line.words
        ]

        bbox = self._bbox(
            all_words
        )

        anchor_consistency = (
            sum(
                self._line_anchor_similarity(
                    lines[i - 1],
                    lines[i],
                    assignment_tolerance,
                )
                for i in range(
                    1,
                    len(lines),
                )
            )
            /
            max(
                1,
                len(lines) - 1,
            )
        )

        confidence = min(
            0.99,

            0.25

            +
            0.30
            *
            min(
                1.0,
                density / 0.75,
            )

            +
            0.30
            *
            anchor_consistency

            +
            0.15
            *
            min(
                1.0,
                populated_rows
                /
                row_count,
            ),
        )

        return TableCandidate(
            table_id=(
                f"p{page_number:03d}"
                f"_t{candidate_index:03d}"
            ),

            page_number=page_number,

            bbox=bbox,

            row_count=row_count,

            column_count=column_count,

            cells=cells,

            confidence=round(
                confidence,
                4,
            ),

            evidence={
                "method":
                    "repeated_horizontal_anchor_detection",

                "density":
                    round(
                        density,
                        4,
                    ),

                "anchor_consistency":
                    round(
                        anchor_consistency,
                        4,
                    ),

                "populated_row_ratio":
                    round(
                        populated_rows
                        /
                        row_count,
                        4,
                    ),

                "column_starts":
                    [
                        round(
                            x,
                            3,
                        )
                        for x in columns
                    ],

                "row_word_counts":
                    [
                        len(line.words)
                        for line in lines
                    ],
            },
        )

    def _detect_tables(
        self,
        page_number: int,
        lines: list[Line],
    ) -> list[TableCandidate]:

        candidates = []

        runs = (
            self._find_table_row_runs(
                lines
            )
        )

        for index, run in enumerate(
            runs,
            start=1,
        ):

            candidate = (
                self._candidate_from_run(
                    page_number,
                    run,
                    index,
                )
            )

            if candidate is not None:
                candidates.append(
                    candidate
                )

        return (
            self._merge_overlapping_tables(
                candidates
            )
        )

    def _merge_overlapping_tables(
        self,
        tables: list[TableCandidate],
    ) -> list[TableCandidate]:

        if len(tables) < 2:
            return tables

        result = []

        for table in sorted(
            tables,
            key=lambda t:
                self._bbox_area(
                    t.bbox
                ),
            reverse=True,
        ):

            duplicate = False

            for existing in result:

                intersection = (
                    self._intersection_area(
                        table.bbox,
                        existing.bbox,
                    )
                )

                smaller = min(
                    self._bbox_area(
                        table.bbox
                    ),
                    self._bbox_area(
                        existing.bbox
                    ),
                )

                if (
                    smaller > 0
                    and
                    (
                        intersection
                        /
                        smaller
                    )
                    > 0.70
                ):

                    duplicate = True
                    break

            if not duplicate:
                result.append(
                    table
                )

        return sorted(
            result,
            key=lambda t: (
                t.bbox["y0"],
                t.bbox["x0"],
            ),
        )

    # ---------------------------------------------------------
    # Public API
    # ---------------------------------------------------------

    def analyze(
        self,
        raw_document: dict[str, Any],
    ) -> StructuredDocument:

        page_results = []

        for page in raw_document.get(
            "pages",
            [],
        ):

            page_number = int(
                page["page_number"]
            )

            words = (
                self._load_words(
                    page
                )
            )

            lines = (
                self._cluster_lines(
                    words
                )
            )

            tables = (
                self._detect_tables(
                    page_number,
                    lines,
                )
            )

            # Remove only words actually assigned to cells. Nearby headings,
            # captions and unassigned words stay available for interpretation.
            table_words = {i for table in tables for cell in table.cells
                           for i in cell.source_word_indexes}
            text_lines = [Line([w for w in line.words if w.index not in table_words])
                          for line in lines]
            regions = self._make_regions(page_number, [line for line in text_lines if line.words])

            # Table candidates are also regions.
            for table in tables:

                regions.append(
                    Region(
                        region_id=(
                            f"{table.table_id}"
                            "_region"
                        ),

                        page_number=
                            page_number,

                        region_type=
                            "table_candidate",

                        bbox=
                            table.bbox,

                        element_indexes=[
                            index
                            for cell
                            in table.cells
                            for index
                            in cell.source_word_indexes
                        ],

                        text="\n".join(
                            cell.text
                            for cell
                            in table.cells
                        ),

                        confidence=
                            table.confidence,

                        evidence=
                            table.evidence,
                    )
                )

            page_results.append(
                PageStructure(
                    warnings=([f"Page {page_number}: OCR may be required; image content has not been interpreted."]
                              if not words or (page.get("needs_ocr_candidate") and page.get("embedded_image_count", 0) > 0) else []),
                    page_number=
                        page_number,

                    regions=
                        regions,

                    tables=
                        tables,
                )
            )

        return StructuredDocument(
            source_schema=
                raw_document.get(
                    "schema_name",
                    "unknown",
                ),

            source_schema_version=
                raw_document.get(
                    "schema_version",
                    "unknown",
                ),

            document=
                raw_document.get(
                    "document",
                    {},
                ),

            pages=
                page_results,
        )


def analyze_file(
    raw_json_path: str | Path,
    output_json_path: str | Path,
    **analyzer_options: Any,
) -> Path:

    raw_path = (
        Path(raw_json_path)
        .expanduser()
        .resolve()
    )

    output_path = (
        Path(output_json_path)
        .expanduser()
        .resolve()
    )

    if not raw_path.exists():
        raise FileNotFoundError(
            f"Raw JSON not found: {raw_path}"
        )

    payload = json.loads(
        raw_path.read_text(
            encoding="utf-8"
        )
    )

    # Validate the Stage-1 contract before indexing into the
    # payload below.  Without this, a malformed JSON file fails later
    # with an unhelpful KeyError or AttributeError.
    raw_document = RawDocument.model_validate(
        payload
    )

    analyzer = StructureAnalyzer(
        **analyzer_options
    )

    structured = analyzer.analyze(
        raw_document.model_dump(
            mode="json"
        )
    )

    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    output_path.write_text(
        json.dumps(
            structured.model_dump(
                mode="json"
            ),
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )

    return output_path
