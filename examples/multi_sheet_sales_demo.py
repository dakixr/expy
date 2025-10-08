from pathlib import Path

import expy as ex


def stat_card(title: str, value: str, delta: str, *, positive: bool = True) -> ex.Node:
    delta_style = ex.text_green if positive else ex.text_red
    return ex.table(
        header=[ex.cell(style=[ex.text_sm, ex.text_gray, ex.text_left])[title]],
        style=[ex.table_bordered, ex.table_compact],
    )[
        [ex.cell(style=[ex.text_2xl, ex.bold, ex.text_black])[value]],
        [ex.cell(style=[ex.text_sm, delta_style])[delta]],
    ]


def summary_section() -> ex.Node:
    headline = ex.row(style=[ex.text_3xl, ex.bold, ex.text_blue])[
        "Q3 Revenue Performance"
    ]

    cards = ex.hstack(
        stat_card("Revenue", "$4.2M", "+14% vs LY", positive=True),
        stat_card("Win Rate", "52%", "+6 pts", positive=True),
        stat_card("Avg. Deal", "$18.9K", "-3% vs LY", positive=False),
        stat_card("Pipeline", "$6.8M", "+$1.1M QoQ", positive=True),
        gap=1,
    )

    regional_performance = ex.table(
        header=[
            ex.cell(style=[ex.text_sm, ex.text_gray, ex.text_left])["Region"],
            ex.cell(style=[ex.text_sm, ex.text_gray, ex.text_right])["GM"],
            ex.cell(style=[ex.text_sm, ex.text_gray, ex.text_right])["Units"],
            ex.cell(style=[ex.text_sm, ex.text_gray, ex.text_right])["YoY"],
        ],
        style=[ex.table_bordered, ex.table_banded, ex.table_compact],
    )[
        [
            "EMEA",
            ex.cell(style=[ex.text_right])["$1.6M"],
            ex.cell(style=[ex.text_right])[1200],
            ex.cell(style=[ex.text_right])["+18%"],
        ],
        [
            "APAC",
            ex.cell(style=[ex.text_right])["$1.1M"],
            ex.cell(style=[ex.text_right])[930],
            ex.cell(style=[ex.text_right])["+9%"],
        ],
        [
            "AMER",
            ex.cell(style=[ex.text_right])["$1.5M"],
            ex.cell(style=[ex.text_right])[1480],
            ex.cell(style=[ex.text_right])["+6%"],
        ],
    ]

    top_opportunities = ex.table(
        header=[
            ex.cell(style=[ex.text_sm, ex.text_gray])["Opportunity"],
            ex.cell(style=[ex.text_sm, ex.text_gray])["Stage"],
            ex.cell(style=[ex.text_sm, ex.text_gray])["Owner"],
            ex.cell(style=[ex.text_sm, ex.text_gray, ex.text_right])["Value"],
        ],
        style=[ex.table_bordered, ex.table_compact],
    )[
        [
            "Atlas Renewals",
            "Negotiation",
            "S. Patel",
            ex.cell(style=[ex.text_right])["$420K"],
        ],
        [
            "Aurora Launch",
            "Proposal",
            "C. Rivers",
            ex.cell(style=[ex.text_right])["$310K"],
        ],
        [
            "Nimbus Edge",
            "Discovery",
            "L. Gomez",
            ex.cell(style=[ex.text_right])["$185K"],
        ],
    ]

    key_updates = ex.table(
        header=["Key Updates"],
        header_style=[ex.text_sm, ex.text_gray],
        style=[ex.table_bordered],
    )[
        [
            ex.cell(style=[ex.wrap])[
                "• APAC backlog cleared; normalization expected by Q4."
            ]
        ],
        [
            ex.cell(style=[ex.wrap])[
                "• Marketing launch for Nimbus Edge driving 23% lift in leads."
            ]
        ],
        [
            ex.cell(style=[ex.wrap])[
                "• Supply constraints eased; lead times back under 4 weeks."
            ]
        ],
    ]

    lower_row = ex.hstack(
        regional_performance,
        top_opportunities,
        key_updates,
        gap=1,
    )

    return ex.vstack(
        headline,
        cards,
        ex.space(),
        lower_row,
        ex.space(),
        ex.row(style=[ex.text_sm, ex.text_gray])["Generated with expy"],
        gap=1,
    )


def raw_data_sheet() -> ex.SheetNode:
    data_table = ex.table(
        header=[
            ex.cell(style=[ex.text_sm, ex.text_gray])["Region"],
            ex.cell(style=[ex.text_sm, ex.text_gray])["Segment"],
            ex.cell(style=[ex.text_sm, ex.text_gray])["Owner"],
            ex.cell(style=[ex.text_sm, ex.text_gray, ex.text_right])["Units"],
            ex.cell(style=[ex.text_sm, ex.text_gray, ex.text_right])["GM"],
        ],
        style=[ex.table_bordered, ex.table_compact],
    )[
        [
            "EMEA",
            "Enterprise",
            "S. Patel",
            ex.cell(style=[ex.text_right])[620],
            ex.cell(style=[ex.text_right])["$820K"],
        ],
        [
            "EMEA",
            "Mid-Market",
            "T. Kato",
            ex.cell(style=[ex.text_right])[580],
            ex.cell(style=[ex.text_right])["$780K"],
        ],
        [
            "APAC",
            "Enterprise",
            "L. Gomez",
            ex.cell(style=[ex.text_right])[410],
            ex.cell(style=[ex.text_right])["$610K"],
        ],
        [
            "APAC",
            "SMB",
            "K. Zhao",
            ex.cell(style=[ex.text_right])[520],
            ex.cell(style=[ex.text_right])["$490K"],
        ],
        [
            "AMER",
            "Enterprise",
            "M. Shaw",
            ex.cell(style=[ex.text_right])[870],
            ex.cell(style=[ex.text_right])["$910K"],
        ],
        [
            "AMER",
            "SMB",
            "C. Rivers",
            ex.cell(style=[ex.text_right])[610],
            ex.cell(style=[ex.text_right])["$590K"],
        ],
    ]

    totals = ex.row()[
        ex.cell(style=[ex.bold, ex.text_gray])["Total"],
        "",
        "",
        ex.cell(style=[ex.bold, ex.text_right])[3610],
        ex.cell(style=[ex.bold, ex.text_right])["$4.2M"],
    ]

    notes = ex.table(
        header=["Notes"],
        header_style=[ex.text_sm, ex.text_gray],
        style=[ex.table_bordered],
    )[
        [
            ex.cell(style=[ex.wrap])[
                "Conversion benchmarks calculated using trailing 90 days."
            ]
        ],
    ]

    return ex.sheet("Raw Data")[
        ex.vstack(
            ex.row(style=[ex.text_lg, ex.bold])["Source Transactions"],
            ex.space(),
            data_table,
            totals,
            ex.space(),
            notes,
            gap=1,
        )
    ]


def pipeline_sheet() -> ex.SheetNode:
    funnel = ex.table(
        header=[
            ex.cell(style=[ex.text_sm, ex.text_gray])["Stage"],
            ex.cell(style=[ex.text_sm, ex.text_gray, ex.text_right])["Deals"],
            ex.cell(style=[ex.text_sm, ex.text_gray, ex.text_right])["Value"],
        ],
        style=[ex.table_bordered, ex.table_compact],
    )[
        [
            "Discovery",
            ex.cell(style=[ex.text_right])[42],
            ex.cell(style=[ex.text_right])["$1.9M"],
        ],
        [
            "Qualification",
            ex.cell(style=[ex.text_right])[33],
            ex.cell(style=[ex.text_right])["$1.4M"],
        ],
        [
            "Proposal",
            ex.cell(style=[ex.text_right])[21],
            ex.cell(style=[ex.text_right])["$1.1M"],
        ],
        [
            "Negotiation",
            ex.cell(style=[ex.text_right])[14],
            ex.cell(style=[ex.text_right])["$1.0M"],
        ],
        [
            "Closed Won",
            ex.cell(style=[ex.text_right])[18],
            ex.cell(style=[ex.text_right])["$1.5M"],
        ],
    ]

    forecast = ex.table(
        header=[
            ex.cell(style=[ex.text_sm, ex.text_gray])["Scenario"],
            ex.cell(style=[ex.text_sm, ex.text_gray, ex.text_right])["Probability"],
            ex.cell(style=[ex.text_sm, ex.text_gray, ex.text_right])["Forecast"],
        ],
        style=[ex.table_bordered, ex.table_compact],
    )[
        [
            "Commit",
            ex.cell(style=[ex.text_right])["75%"],
            ex.cell(style=[ex.text_right])["$3.2M"],
        ],
        [
            "Best",
            ex.cell(style=[ex.text_right])["50%"],
            ex.cell(style=[ex.text_right])["$4.5M"],
        ],
        [
            "Upside",
            ex.cell(style=[ex.text_right])["25%"],
            ex.cell(style=[ex.text_right])["$6.1M"],
        ],
    ]

    team_heatmap = ex.table(
        header=[
            ex.cell(style=[ex.text_sm, ex.text_gray])["Owner"],
            ex.cell(style=[ex.text_sm, ex.text_gray, ex.text_right])["Active Deals"],
            ex.cell(style=[ex.text_sm, ex.text_gray, ex.text_right])["Win Rate"],
        ],
        style=[ex.table_bordered, ex.table_banded, ex.table_compact],
    )[
        [
            "S. Patel",
            ex.cell(style=[ex.text_right])[18],
            ex.cell(style=[ex.text_right])["64%"],
        ],
        [
            "C. Rivers",
            ex.cell(style=[ex.text_right])[15],
            ex.cell(style=[ex.text_right])["58%"],
        ],
        [
            "L. Gomez",
            ex.cell(style=[ex.text_right])[12],
            ex.cell(style=[ex.text_right])["54%"],
        ],
        [
            "T. Kato",
            ex.cell(style=[ex.text_right])[11],
            ex.cell(style=[ex.text_right])["49%"],
        ],
    ]

    return ex.sheet("Pipeline")[
        ex.vstack(
            ex.row(style=[ex.text_lg, ex.bold])["Pipeline & Forecast"],
            ex.space(),
            ex.hstack(funnel, forecast, gap=2),
            ex.space(),
            team_heatmap,
            gap=1,
        )
    ]


def glossary_sheet() -> ex.SheetNode:
    utility_grid = ex.table(
        header=["Utility", "Description", "Example"],
        header_style=[ex.text_sm, ex.text_gray],
        style=[ex.table_bordered, ex.table_compact],
    )[
        ["text_blue", "Accent headline color", "Q3 Revenue"],
        ["bg_success", "Positive badge background", "+14% Revenue"],
        ["wrap", "Wrap long text inside cells", "Supply constraints eased..."],
        ["number_precision", "Two-decimal numeric format", "42100.00"],
    ]

    return ex.sheet("Glossary")[
        ex.vstack(
            ex.row(style=[ex.text_lg, ex.bold])["Utility Cheatsheet"],
            ex.space(),
            utility_grid,
            gap=1,
        )
    ]


def build_sample_workbook() -> ex.Workbook:
    summary_sheet = ex.sheet("Summary")[summary_section()]

    workbook = ex.workbook("Sales Demo")[
        summary_sheet,
        raw_data_sheet(),
        pipeline_sheet(),
        glossary_sheet(),
    ]

    return workbook


def main() -> None:
    output_path = Path("multi-sheet-sales-demo-output.xlsx")
    workbook = build_sample_workbook()
    workbook.save(output_path)
    print(f"Saved workbook to {output_path.resolve()}")


if __name__ == "__main__":
    main()
