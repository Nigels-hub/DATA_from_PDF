# ─────────────────────────────────────────────────────────────────────────────
# Shiny app – Interaction P-Values by Subgroup
#
# Reads all "*by Subgroups*.xlsx" files from extracted_tables/, parses
# subgroup categories and their interaction p-values, and shows an
# interactive forest-style plot.
#
# Run from the DATA_from_PDF directory:
#   Rscript -e "shiny::runApp('interaction_pvalues_app.R')"
# ─────────────────────────────────────────────────────────────────────────────

library(shiny)
library(plotly)
library(readxl)
library(dplyr)

# ── Parsing ──────────────────────────────────────────────────────────────────

extract_outcome <- function(title) {
  dplyr::case_when(
    grepl("Overall Survival",    title, ignore.case = TRUE) ~ "Overall Survival (OS)",
    grepl("Progression-Free.*IRC", title, ignore.case = TRUE) ~ "PFS \u2013 IRC",
    grepl("Progression-Free.*INV", title, ignore.case = TRUE) ~ "PFS \u2013 INV",
    TRUE ~ "Other"
  )
}

parse_subgroup_file <- function(filepath) {
  raw_tbl <- tryCatch(
    suppressMessages(
      readxl::read_excel(filepath, col_names = FALSE, .name_repair = "minimal")
    ),
    error = function(e) NULL
  )
  if (is.null(raw_tbl) || nrow(raw_tbl) == 0L) return(NULL)

  # Work with a plain character matrix – simplest and most robust
  mat <- matrix(as.character(as.matrix(raw_tbl)), nrow = nrow(raw_tbl))
  mat[mat == "NA"] <- NA_character_
  for (j in seq_len(ncol(mat))) mat[, j] <- trimws(mat[, j])
  mat[mat == ""] <- NA_character_
  nr <- nrow(mat); nc <- ncol(mat)

  cv <- function(r, c) {
    if (r < 1L || r > nr || c < 1L || c > nc) return(NA_character_)
    mat[r, c]
  }

  # Title = cell (1,1)
  title <- cv(1L, 1L)
  if (is.na(title)) return(NULL)

  # Find the row where col-1 == "Subgroup"
  subg_row <- NA_integer_
  for (r in seq_len(nr)) {
    if (identical(cv(r, 1L), "Subgroup")) { subg_row <- r; break }
  }
  if (is.na(subg_row)) return(NULL)

  # Find the interaction p-value column:
  # scan the header block (rows 1..subg_row) for "Interaction" text
  int_col <- NA_integer_
  for (c in seq(nc, 1L, by = -1L)) {
    for (r in seq_len(subg_row)) {
      v <- cv(r, c)
      if (!is.na(v) && grepl("interaction", v, ignore.case = TRUE)) {
        int_col <- c; break
      }
    }
    if (!is.na(int_col)) break
  }
  if (is.na(int_col)) return(NULL)

  # ── Row-by-row parse ─────────────────────────────────────────────────────
  results   <- list()
  cat_parts <- character(0)   # accumulates multi-line category names
  sub_vals  <- character(0)   # subgroup value labels within current category
  pval      <- NA_real_

  flush_record <- function() {
    if (!is.na(pval) && length(cat_parts) > 0L) {
      results[[length(results) + 1L]] <<- data.frame(
        title           = title,
        outcome         = extract_outcome(title),
        category        = paste(cat_parts, collapse = " "),
        subgroup_values = paste(sub_vals,  collapse = ", "),
        interaction_p   = pval,
        stringsAsFactors = FALSE
      )
    }
  }

  for (i in seq(subg_row + 1L, nr)) {
    c1 <- cv(i, 1L)
    c2 <- cv(i, 2L)

    if (is.na(c1)) next

    # Stop at the footnotes / abbreviations section
    if (grepl("^(Notes|Abbreviations)", c1, ignore.case = TRUE)) break

    if (is.na(c2)) {
      # ── Category-header row (or continuation of a multi-line name) ────────
      if (length(sub_vals) > 0L) {
        # Previous category is complete – save it and start fresh
        flush_record()
        cat_parts <- c1
        sub_vals  <- character(0)
        pval      <- NA_real_
      } else {
        # Still building the category name (e.g. "Number of Organs with" /
        # "Metastatic Sites" or "Prior Gastrectomy" / "(total or partial)")
        cat_parts <- c(cat_parts, c1)
      }
    } else {
      # ── Data row ──────────────────────────────────────────────────────────
      sub_vals <- c(sub_vals, c1)

      # Interaction p-value is on the FIRST data row of each category only
      if (is.na(pval)) {
        raw_p <- cv(i, int_col)
        if (!is.na(raw_p)) {
          p_num <- suppressWarnings(as.numeric(raw_p))
          if (!is.na(p_num)) pval <- p_num
        }
      }
    }
  }
  flush_record()   # save the last category

  if (length(results) == 0L) return(NULL)
  dplyr::bind_rows(results)
}

load_all_data <- function(tables_dir = "extracted_tables") {
  files <- list.files(
    tables_dir,
    pattern     = "Subgroups.*\\.xlsx$",
    full.names  = TRUE,
    ignore.case = TRUE
  )
  if (length(files) == 0L) return(data.frame())
  dplyr::bind_rows(lapply(files, parse_subgroup_file))
}

# ── UI ───────────────────────────────────────────────────────────────────────

ui <- fluidPage(
  tags$head(tags$style(HTML("
    body         { font-family: 'Segoe UI', Arial, sans-serif; background:#f5f5f5; }
    .well        { background:white; border:1px solid #ddd; border-radius:6px; }
    h4           { color:#2c3e50; margin-top:0; }
    .plot-wrap   { background:white; border-radius:6px;
                   box-shadow:0 1px 4px rgba(0,0,0,.1); padding:12px; }
  "))),

  titlePanel(
    div(style = "color:#2c3e50;",
        "Subgroup Interaction P-Values",
        tags$small(style = "font-size:14px; color:#7f8c8d; margin-left:12px;",
                   "Zolbetuximab + mFOLFOX6 vs. Placebo + mFOLFOX6"))
  ),

  sidebarLayout(
    sidebarPanel(
      width = 3,
      h4("Outcomes"),
      checkboxGroupInput("sel_outcomes", label = NULL, choices = character(0)),
      hr(),
      h4("Significance threshold"),
      sliderInput("threshold",
                  label = "Highlight p \u2264",
                  min = 0.01, max = 1.0, value = 0.05, step = 0.01),
      hr(),
      tags$small(style = "color:#666;",
        tags$b("Solid / coloured"), " points: p \u2264 threshold.",  tags$br(),
        tags$b("Grey"), " points: p > threshold.",                  tags$br(), tags$br(),
        "Hover over any point to see the full details.")
    ),

    mainPanel(
      width = 9,
      div(class = "plot-wrap",
          plotlyOutput("forest_plot", height = "auto"))
    )
  )
)

# ── Server ───────────────────────────────────────────────────────────────────

server <- function(input, output, session) {

  # Load once at startup
  all_data <- reactive({ load_all_data("extracted_tables") })

  # Populate outcome checkboxes
  observe({
    d <- all_data()
    if (nrow(d) == 0L) return()
    ocs <- sort(unique(d$outcome))
    updateCheckboxGroupInput(session, "sel_outcomes",
                             choices = ocs, selected = ocs)
  })

  # Filtered dataset
  fdata <- reactive({
    d <- all_data()
    if (nrow(d) == 0L || is.null(input$sel_outcomes)) return(d)
    dplyr::filter(d, outcome %in% input$sel_outcomes)
  })

  output$forest_plot <- renderPlotly({
    d <- fdata()
    if (nrow(d) == 0L) {
      return(
        plot_ly() |>
          layout(title = "No data found – check extracted_tables/ directory")
      )
    }

    thresh <- input$threshold

    # Y-axis: unique categories, reversed so first category is at the top
    cats  <- rev(unique(d$category))
    cat_y <- setNames(seq_along(cats), cats)

    oc_order   <- c("Overall Survival (OS)", "PFS \u2013 IRC", "PFS \u2013 INV")
    oc_present <- intersect(oc_order, unique(d$outcome))
    n_oc       <- length(oc_present)
    offsets    <- setNames(
      seq(-0.22, 0.22, length.out = max(n_oc, 1L)),
      oc_present
    )

    palette <- c(
      "Overall Survival (OS)" = "#c0392b",
      "PFS \u2013 IRC"        = "#2471a3",
      "PFS \u2013 INV"        = "#1e8449"
    )

    # Dynamic height so all categories are visible without scrolling
    plot_height <- max(500L, length(cats) * 52L + 160L)

    fig <- plot_ly(height = plot_height)

    for (oc in oc_present) {
      dd  <- dplyr::filter(d, outcome == oc)
      yy  <- unname(cat_y[dd$category]) + offsets[[oc]]
      col <- palette[[oc]]

      # Points significant vs non-significant get different colours
      pt_color <- ifelse(dd$interaction_p <= thresh, col, "rgba(190,190,190,0.45)")
      bd_color <- ifelse(dd$interaction_p <= thresh, col, "rgba(150,150,150,0.55)")

      hover_text <- paste0(
        "<b>", oc, "</b><br>",
        "Subgroup category: <b>", dd$category, "</b><br>",
        "Values compared: ",      dd$subgroup_values, "<br>",
        "p(interaction) = <b>",  sprintf("%.4f", dd$interaction_p), "</b>",
        ifelse(dd$interaction_p <= thresh,
               paste0(" <span style='color:", col, ";'>\u25cf significant</span>"),
               ""),
        "<br><i style='font-size:11px;color:#999;'>", dd$title, "</i>"
      )

      fig <- add_trace(fig,
        x         = dd$interaction_p,
        y         = yy,
        type      = "scatter",
        mode      = "markers",
        name      = oc,
        legendgroup = oc,
        marker    = list(
          size   = 14,
          color  = pt_color,
          symbol = "circle",
          line   = list(width = 1.8, color = bd_color)
        ),
        text      = hover_text,
        hoverinfo = "text"
      )
    }

    # Vertical threshold line
    fig <- add_trace(fig,
      x           = c(thresh, thresh),
      y           = c(0.5, length(cats) + 0.5),
      type        = "scatter",
      mode        = "lines",
      line        = list(color = "#c0392b", width = 1.5, dash = "dash"),
      name        = sprintf("p = %.2f threshold", thresh),
      legendgroup = "threshold",
      hoverinfo   = "skip"
    )

    # Dynamic height so all categories are visible
    plot_height <- max(500L, length(cats) * 52L + 160L)

    layout(fig,
      title  = list(
        text = "Interaction P-Values by Subgroup",
        font = list(size = 15, color = "#2c3e50"),
        x    = 0.02
      ),
      xaxis = list(
        title     = "Interaction P-Value",
        range     = c(-0.03, 1.05),
        zeroline  = FALSE,
        gridcolor = "#ececec",
        tickformat = ".2f"
      ),
      yaxis = list(
        title      = "",
        tickvals   = unname(cat_y),
        ticktext   = names(cat_y),
        automargin = TRUE,
        tickfont   = list(size = 11),
        gridcolor  = "#ececec"
      ),
      legend = list(
        orientation = "h",
        x = 0, y = -0.12,
        font = list(size = 12)
      ),
      hovermode     = "closest",
      plot_bgcolor  = "white",
      paper_bgcolor = "white",
      margin        = list(l = 260, r = 30, t = 60, b = 110)
    )
  })
}

shinyApp(ui, server)
