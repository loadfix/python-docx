"""Unit test suite for the docx.parts.chart module."""

from __future__ import annotations

import pytest

from docx.chart import WD_CHART_TYPE
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import PartFactory
from docx.oxml.chart import CT_ChartSpace
from docx.package import Package
from docx.parts.chart import ChartPart

from ..unitutil.mock import FixtureRequest, Mock, instance_mock, method_mock


class DescribeChartPart:
    """Unit test suite for `docx.parts.chart.ChartPart`."""

    def it_is_used_by_the_part_loader_to_construct_a_chart_part(
        self, package_: Mock, ChartPart_load_: Mock, chart_part_: Mock
    ):
        partname = PackURI("/word/charts/chart1.xml")
        content_type = CT.DML_CHART
        reltype = RT.CHART
        blob = (
            b'<c:chartSpace xmlns:c="http://schemas.openxmlformats.org'
            b'/drawingml/2006/chart"/>'
        )
        ChartPart_load_.return_value = chart_part_

        part = PartFactory(partname, content_type, reltype, blob, package_)

        ChartPart_load_.assert_called_once_with(
            partname, content_type, blob, package_
        )
        assert part is chart_part_

    def it_constructs_a_new_bar_chart_part(self):
        package = Package()
        part = ChartPart.new(
            package,
            WD_CHART_TYPE.BAR,
            ["a", "b", "c"],
            {"Series 1": [1.0, 2.0, 3.0]},
        )
        assert isinstance(part, ChartPart)
        assert part.content_type == CT.DML_CHART
        assert part.partname.startswith("/word/charts/chart")
        assert part.partname.endswith(".xml")
        assert isinstance(part.chartSpace, CT_ChartSpace)
        # -- round-trip the chart data via the element tree --
        chart = part.chartSpace.chart
        assert chart is not None
        plotArea = chart.plotArea
        assert plotArea is not None
        kind = plotArea.chart_kind_element
        assert kind is not None
        assert kind.tag.endswith("barChart")

    @pytest.mark.parametrize(
        ("chart_type", "expected_tag"),
        [
            (WD_CHART_TYPE.BAR, "barChart"),
            (WD_CHART_TYPE.BAR_STACKED, "barChart"),
            (WD_CHART_TYPE.COLUMN, "barChart"),
            (WD_CHART_TYPE.COLUMN_STACKED, "barChart"),
            (WD_CHART_TYPE.LINE, "lineChart"),
            (WD_CHART_TYPE.PIE, "pieChart"),
        ],
    )
    def it_constructs_new_parts_for_each_supported_kind(
        self, chart_type: WD_CHART_TYPE, expected_tag: str
    ):
        package = Package()
        part = ChartPart.new(
            package, chart_type, ["a", "b"], {"S": [1.0, 2.0]}
        )
        kind = part.chartSpace.chart.plotArea.chart_kind_element  # pyright: ignore
        assert kind is not None
        assert kind.tag.endswith(expected_tag)

    def it_rejects_mismatched_series_length(self):
        package = Package()
        with pytest.raises(ValueError, match="has 2 values but 3 categories"):
            ChartPart.new(
                package,
                WD_CHART_TYPE.BAR,
                ["a", "b", "c"],
                {"S": [1.0, 2.0]},
            )

    def it_rejects_unsupported_chart_type_for_create(self):
        package = Package()
        with pytest.raises(ValueError, match="unsupported chart_type"):
            ChartPart.new(
                package,
                WD_CHART_TYPE.SCATTER,
                ["a"],
                {"S": [1.0]},
            )

    def it_generates_sequential_partnames(self):
        package = Package()
        part1 = ChartPart.new(
            package, WD_CHART_TYPE.BAR, ["a"], {"S": [1.0]}
        )
        package.relate_to(part1, RT.CHART)
        part2 = ChartPart.new(
            package, WD_CHART_TYPE.BAR, ["a"], {"S": [1.0]}
        )
        # -- partnames must be distinct so they don't collide in the package --
        assert part1.partname != part2.partname

    # -- fixtures ---------------------------------------------------------

    @pytest.fixture
    def chart_part_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, ChartPart)

    @pytest.fixture
    def ChartPart_load_(self, request: FixtureRequest) -> Mock:
        return method_mock(request, ChartPart, "load", autospec=False)

    @pytest.fixture
    def package_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, Package)
