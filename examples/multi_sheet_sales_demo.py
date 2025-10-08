from pathlib import Path

import xpyxl as x


def stat_card(title: str, value: str, delta: str, *, positive: bool = True) -> x.Node:
    delta_style = x.text_green if positive else x.text_red
    return x.table(
        header=[x.cell(style=[x.text_sm, x.text_gray, x.text_left])[title]],
        style=[x.table_bordered, x.table_compact],
    )[
        [x.cell(style=[x.text_2xl, x.bold, x.text_black])[value]],
        [x.cell(style=[x.text_sm, delta_style])[delta]],
    ]


def summary_section() -> x.Node:
    headline = x.row(style=[x.text_3xl, x.bold, x.text_blue])["Q3 Revenue Performance"]

    cards = x.hstack(
        stat_card("Revenue", "$4.2M", "+14% vs LY", positive=True),
        stat_card("Win Rate", "52%", "+6 pts", positive=True),
        stat_card("Avg. Deal", "$18.9K", "-3% vs LY", positive=False),
        stat_card("Pipeline", "$6.8M", "+$1.1M QoQ", positive=True),
        gap=1,
    )

    regional_performance = x.table(
        header=[
            x.cell(style=[x.text_sm, x.text_gray, x.text_left])["Region"],
            x.cell(style=[x.text_sm, x.text_gray, x.text_right])["GM"],
            x.cell(style=[x.text_sm, x.text_gray, x.text_right])["Units"],
            x.cell(style=[x.text_sm, x.text_gray, x.text_right])["YoY"],
        ],
        style=[x.table_bordered, x.table_banded, x.table_compact],
    )[
        [
            "EMEA",
            x.cell(style=[x.text_right])["$1.6M"],
            x.cell(style=[x.text_right])[1200],
            x.cell(style=[x.text_right])["+18%"],
        ],
        [
            "APAC",
            x.cell(style=[x.text_right])["$1.1M"],
            x.cell(style=[x.text_right])[930],
            x.cell(style=[x.text_right])["+9%"],
        ],
        [
            "AMER",
            x.cell(style=[x.text_right])["$1.5M"],
            x.cell(style=[x.text_right])[1480],
            x.cell(style=[x.text_right])["+6%"],
        ],
    ]

    top_opportunities = x.table(
        header=[
            x.cell(style=[x.text_sm, x.text_gray])["Opportunity"],
            x.cell(style=[x.text_sm, x.text_gray])["Stage"],
            x.cell(style=[x.text_sm, x.text_gray])["Owner"],
            x.cell(style=[x.text_sm, x.text_gray, x.text_right])["Value"],
        ],
        style=[x.table_bordered, x.table_compact],
    )[
        [
            "Atlas Renewals",
            "Negotiation",
            "S. Patel",
            x.cell(style=[x.text_right])["$420K"],
        ],
        [
            "Aurora Launch",
            "Proposal",
            "C. Rivers",
            x.cell(style=[x.text_right])["$310K"],
        ],
        [
            "Nimbus Edge",
            "Discovery",
            "L. Gomez",
            x.cell(style=[x.text_right])["$185K"],
        ],
    ]

    key_updates = x.table(
        header=["Key Updates"],
        header_style=[x.text_sm, x.text_gray],
        style=[x.table_bordered],
    )[
        [
            x.cell(style=[x.wrap])[
                "• APAC backlog cleared; normalization expected by Q4."
            ]
        ],
        [
            x.cell(style=[x.wrap])[
                "• Marketing launch for Nimbus Edge driving 23% lift in leads."
            ]
        ],
        [
            x.cell(style=[x.wrap])[
                "• Supply constraints eased; lead times back under 4 weeks."
            ]
        ],
    ]

    lower_row = x.hstack(
        regional_performance,
        top_opportunities,
        key_updates,
        gap=1,
    )

    return x.vstack(
        headline,
        cards,
        x.space(),
        lower_row,
        x.space(),
        x.row(style=[x.text_sm, x.text_gray])["Generated with xsxpy"],
        gap=1,
    )


def raw_data_sheet() -> x.SheetNode:
    data_table = x.table(
        header=[
            x.cell(style=[x.text_sm, x.text_gray])["Region"],
            x.cell(style=[x.text_sm, x.text_gray])["Segment"],
            x.cell(style=[x.text_sm, x.text_gray])["Owner"],
            x.cell(style=[x.text_sm, x.text_gray, x.text_right])["Units"],
            x.cell(style=[x.text_sm, x.text_gray, x.text_right])["GM"],
        ],
        style=[x.table_bordered, x.table_compact],
    )[
        [
            "EMEA",
            "Enterprise",
            "S. Patel",
            x.cell(style=[x.text_right])[620],
            x.cell(style=[x.text_right])["$820K"],
        ],
        [
            "EMEA",
            "Mid-Market",
            "T. Kato",
            x.cell(style=[x.text_right])[580],
            x.cell(style=[x.text_right])["$780K"],
        ],
        [
            "APAC",
            "Enterprise",
            "L. Gomez",
            x.cell(style=[x.text_right])[410],
            x.cell(style=[x.text_right])["$610K"],
        ],
        [
            "APAC",
            "SMB",
            "K. Zhao",
            x.cell(style=[x.text_right])[520],
            x.cell(style=[x.text_right])["$490K"],
        ],
        [
            "AMER",
            "Enterprise",
            "M. Shaw",
            x.cell(style=[x.text_right])[870],
            x.cell(style=[x.text_right])["$910K"],
        ],
        [
            "AMER",
            "SMB",
            "C. Rivers",
            x.cell(style=[x.text_right])[610],
            x.cell(style=[x.text_right])["$590K"],
        ],
    ]

    totals = x.row()[
        x.cell(style=[x.bold, x.text_gray])["Total"],
        "",
        "",
        x.cell(style=[x.bold, x.text_right])[3610],
        x.cell(style=[x.bold, x.text_right])["$4.2M"],
    ]

    notes = x.table(
        header=["Notes"],
        header_style=[x.text_sm, x.text_gray],
        style=[x.table_bordered],
    )[
        [
            x.cell(style=[x.wrap])[
                "Conversion benchmarks calculated using trailing 90 days."
            ]
        ],
    ]

    return x.sheet("Raw Data")[
        x.vstack(
            x.row(style=[x.text_lg, x.bold])["Source Transactions"],
            x.space(),
            data_table,
            totals,
            x.space(),
            notes,
            gap=1,
        )
    ]


def pipeline_sheet() -> x.SheetNode:
    funnel = x.table(
        header=[
            x.cell(style=[x.text_sm, x.text_gray])["Stage"],
            x.cell(style=[x.text_sm, x.text_gray, x.text_right])["Deals"],
            x.cell(style=[x.text_sm, x.text_gray, x.text_right])["Value"],
        ],
        style=[x.table_bordered, x.table_compact],
    )[
        [
            "Discovery",
            x.cell(style=[x.text_right])[42],
            x.cell(style=[x.text_right])["$1.9M"],
        ],
        [
            "Qualification",
            x.cell(style=[x.text_right])[33],
            x.cell(style=[x.text_right])["$1.4M"],
        ],
        [
            "Proposal",
            x.cell(style=[x.text_right])[21],
            x.cell(style=[x.text_right])["$1.1M"],
        ],
        [
            "Negotiation",
            x.cell(style=[x.text_right])[14],
            x.cell(style=[x.text_right])["$1.0M"],
        ],
        [
            "Closed Won",
            x.cell(style=[x.text_right])[18],
            x.cell(style=[x.text_right])["$1.5M"],
        ],
    ]

    forecast = x.table(
        header=[
            x.cell(style=[x.text_sm, x.text_gray])["Scenario"],
            x.cell(style=[x.text_sm, x.text_gray, x.text_right])["Probability"],
            x.cell(style=[x.text_sm, x.text_gray, x.text_right])["Forecast"],
        ],
        style=[x.table_bordered, x.table_compact],
    )[
        [
            "Commit",
            x.cell(style=[x.text_right])["75%"],
            x.cell(style=[x.text_right])["$3.2M"],
        ],
        [
            "Best",
            x.cell(style=[x.text_right])["50%"],
            x.cell(style=[x.text_right])["$4.5M"],
        ],
        [
            "Upside",
            x.cell(style=[x.text_right])["25%"],
            x.cell(style=[x.text_right])["$6.1M"],
        ],
    ]

    team_heatmap = x.table(
        header=[
            x.cell(style=[x.text_sm, x.text_gray])["Owner"],
            x.cell(style=[x.text_sm, x.text_gray, x.text_right])["Active Deals"],
            x.cell(style=[x.text_sm, x.text_gray, x.text_right])["Win Rate"],
        ],
        style=[x.table_bordered, x.table_banded, x.table_compact],
    )[
        [
            "S. Patel",
            x.cell(style=[x.text_right])[18],
            x.cell(style=[x.text_right])["64%"],
        ],
        [
            "C. Rivers",
            x.cell(style=[x.text_right])[15],
            x.cell(style=[x.text_right])["58%"],
        ],
        [
            "L. Gomez",
            x.cell(style=[x.text_right])[12],
            x.cell(style=[x.text_right])["54%"],
        ],
        [
            "T. Kato",
            x.cell(style=[x.text_right])[11],
            x.cell(style=[x.text_right])["49%"],
        ],
    ]

    return x.sheet("Pipeline")[
        x.vstack(
            x.row(style=[x.text_lg, x.bold])["Pipeline & Forecast"],
            x.space(),
            x.hstack(funnel, forecast, gap=2),
            x.space(),
            team_heatmap,
            gap=1,
        )
    ]


def glossary_sheet() -> x.SheetNode:
    utility_grid = x.table(
        header=["Utility", "Description", "Example"],
        header_style=[x.text_sm, x.text_gray],
        style=[x.table_bordered, x.table_compact],
    )[
        ["text_blue", "Accent headline color", "Q3 Revenue"],
        ["bg_success", "Positive badge background", "+14% Revenue"],
        ["wrap", "Wrap long text inside cells", "Supply constraints eased..."],
        ["number_precision", "Two-decimal numeric format", "42100.00"],
    ]

    return x.sheet("Glossary")[
        x.vstack(
            x.row(style=[x.text_lg, x.bold])["Utility Cheatsheet"],
            x.space(),
            utility_grid,
            gap=1,
        )
    ]


def build_sample_workbook() -> x.Workbook:
    summary_sheet = x.sheet("Summary")[summary_section()]

    workbook = x.workbook("Sales Demo")[
        summary_sheet,
        raw_data_sheet(),
        pipeline_sheet(),
        glossary_sheet(),
    ]

    return workbook


def main() -> None:
    output_path = Path("multi-sheet-sales-demo-output.xsx")
    workbook = build_sample_workbook()
    workbook.save(output_path)
    print(f"Saved workbook to {output_path.resolve()}")


if __name__ == "__main__":
    main()
