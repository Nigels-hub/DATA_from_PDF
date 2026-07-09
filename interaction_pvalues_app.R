# ─────────────────────────────────────────────────────────────────────────────
# Shiny app – Interaction P-Values by Subgroup
#
# Reads subgroup_interaction_pvalues.csv (produced by prepare_subgroup_data.py)
# and shows an interactive forest-style plot with filters for Domain and
# Endpoint.
#
# Run from the DATA_from_PDF directory:
#   Rscript -e "shiny::runApp('interaction_pvalues_app.R')"
# ─────────────────────────────────────────────────────────────────────────────

pacman::p_load(shiny, plotly, dplyr, readr, DT)

# ── Data loading ──────────────────────────────────────────────────────────────

load_data <- function(csv_path = "subgroup_interaction_pvalues.csv") {
  if (!file.exists(csv_path)) {
    return(data.frame(
      title           = character(0),
      table_number    = integer(0),
      domain          = character(0),
      endpoint_short  = character(0),
      category        = character(0),
      subgroup_values = character(0),
      interaction_p   = numeric(0)
    ))
  }
  d <- readr::read_csv(
    csv_path,
    col_types = readr::cols(
      title           = readr::col_character(),
      table_number    = readr::col_integer(),
      domain          = readr::col_character(),
      endpoint_short  = readr::col_character(),
      category        = readr::col_character(),
      subgroup_values = readr::col_character(),
      interaction_p   = readr::col_double()
    ),
    show_col_types = FALSE
  )
  d$domain <- factor(d$domain, levels = c("Efficacy", "Safety", "PRO", "Unknown"))
  d
}

# kept only for backward compatibility – no longer called internally
parse_subgroup_file <- function(filepath) { NULL }

# ── UI ───────────────────────────────────────────────────────────────────────

ui <- fluidPage(
  tags$head(tags$style(HTML("
    body            { font-family: 'Segoe UI', Arial, sans-serif; background:#f5f5f5; }
    .well           { background:white; border:1px solid #ddd; border-radius:6px; }
    h4              { color:#2c3e50; margin-top:0; }
    .plot-wrap      { background:white; border-radius:6px;
                      box-shadow:0 1px 4px rgba(0,0,0,.1); padding:12px; }
    .details-wrap   { background:white; border-radius:6px;
                      box-shadow:0 1px 4px rgba(0,0,0,.1); padding:14px; margin-top:14px; }
    .sidebar-section { margin-bottom: 14px; }
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

      # ── Domain filter ─────────────────────────────────────────────────────
      div(class = "sidebar-section",
        h4("Domain"),
        checkboxGroupInput(
          "sel_domain",
          label    = NULL,
          choices  = c("Efficacy", "Safety", "PRO"),
          selected = c("Efficacy", "Safety", "PRO")
        )
      ),

      hr(),

      # ── Endpoint filter (updates dynamically based on domain) ─────────────
      div(class = "sidebar-section",
        h4("Endpoints"),
        actionLink("select_all_ep",   "Select all"),
        " | ",
        actionLink("deselect_all_ep", "Deselect all"),
        tags$br(), tags$br(),
        checkboxGroupInput(
          "sel_endpoints",
          label   = NULL,
          choices = character(0)
        )
      ),

      hr(),

      # ── Significance threshold ────────────────────────────────────────────
      div(class = "sidebar-section",
        h4("Significance threshold"),
        sliderInput("threshold",
                    label = "Highlight p \u2264",
                    min = 0.01, max = 1.0, value = 0.05, step = 0.01)
      ),

      hr(),
      tags$small(style = "color:#666;",
        tags$b("Solid / coloured"), " points: p \u2264 threshold.", tags$br(),
        tags$b("Grey"),             " points: p > threshold.",      tags$br(), tags$br(),
        "Hover for details.", tags$br(),
        "Double-click a point to add its table to the details panel."
      )
    ),

    mainPanel(
      width = 9,

      # ── Plot ──────────────────────────────────────────────────────────────
      div(class = "plot-wrap",
          plotlyOutput("forest_plot", height = "auto")),

      # ── Details panel ─────────────────────────────────────────────────────
      div(class = "details-wrap",
        fluidRow(
          column(6,
            h4(style = "margin-bottom:4px;", "Significant Subgroup Details")
          ),
          column(6,
            div(style = "text-align:right;",
              radioButtons("detail_mode", label = NULL,
                choices  = c("Show all significant" = "all",
                             "Show selected"        = "selected"),
                selected = "all", inline = TRUE),
              conditionalPanel(
                condition = "input.detail_mode === 'selected'",
                actionButton("clear_sel", "Clear selection",
                             class = "btn btn-sm btn-default",
                             style = "margin-top:2px;")
              )
            )
          )
        ),
        tags$p(style = "color:#888; font-size:12px; margin:4px 0 8px 0;",
          tags$b("Show all significant:"),
          " all rows from every table that contains at least one p \u2264 threshold.", tags$br(),
          tags$b("Show selected:"),
          " double-click any point on the plot to load its full source table."
        ),
        DT::dataTableOutput("detail_table")
      )
    )
  )
)

# ── Server ───────────────────────────────────────────────────────────────────

server <- function(input, output, session) {

  # Load once at startup
  all_data <- reactive({ load_data("subgroup_interaction_pvalues.csv") })

  # Endpoints available within the selected domains
  available_endpoints <- reactive({
    d <- all_data()
    if (nrow(d) == 0L || is.null(input$sel_domain)) return(character(0))
    d |>
      dplyr::filter(domain %in% input$sel_domain) |>
      dplyr::pull(endpoint_short) |>
      unique() |>
      sort()
  })

  # Keep endpoint checkboxes in sync with domain selection
  observe({
    eps  <- available_endpoints()
    prev <- isolate(input$sel_endpoints)
    still_valid <- intersect(prev, eps)
    selected    <- if (length(still_valid) > 0L) still_valid else eps
    updateCheckboxGroupInput(session, "sel_endpoints",
                             choices  = eps,
                             selected = selected)
  })

  # Select / deselect all endpoints
  observeEvent(input$select_all_ep, {
    updateCheckboxGroupInput(session, "sel_endpoints",
                             selected = available_endpoints())
  })
  observeEvent(input$deselect_all_ep, {
    updateCheckboxGroupInput(session, "sel_endpoints", selected = character(0))
  })

  # Filtered dataset
  fdata <- reactive({
    d <- all_data()
    if (nrow(d) == 0L) return(d)
    req(input$sel_domain, input$sel_endpoints)
    d |>
      dplyr::filter(
        domain         %in% input$sel_domain,
        endpoint_short %in% input$sel_endpoints
      )
  })

  # ── Track table selections from double-clicks ─────────────────────────────

  selected_tables <- reactiveVal(character(0))

  observeEvent(input$clear_sel, {
    selected_tables(character(0))
  })

  # Reset selection when switching back to "all" mode
  observeEvent(input$detail_mode, {
    if (input$detail_mode == "all") selected_tables(character(0))
  })

  observeEvent(event_data("plotly_doubleclick", source = "forest"), {
    evt <- event_data("plotly_doubleclick", source = "forest")
    if (is.null(evt) || is.null(evt$customdata)) return()
    key     <- as.character(evt$customdata)
    current <- selected_tables()
    # Toggle: double-clicking again removes it
    if (key %in% current) {
      selected_tables(setdiff(current, key))
    } else {
      selected_tables(c(current, key))
    }
  })

  # ── Plot ──────────────────────────────────────────────────────────────────

  output$forest_plot <- renderPlotly({
    d <- fdata()

    if (nrow(d) == 0L) {
      return(
        plot_ly() |>
          layout(title = "No data – run prepare_subgroup_data.py first, or adjust filters")
      )
    }

    thresh <- input$threshold

    # Y-axis: unique subgroup categories, reversed so first is at top
    cats  <- rev(unique(d$category))
    cat_y <- setNames(seq_along(cats), cats)

    ep_order <- sort(unique(d$endpoint_short))
    n_ep     <- length(ep_order)

    # Cycle through a qualitative colour palette
    base_colours <- c(
      "#c0392b", "#2471a3", "#1e8449", "#d35400", "#8e44ad",
      "#16a085", "#795548", "#7f8c8d", "#2ecc71", "#e74c3c",
      "#3498db", "#9b59b6", "#f39c12", "#1abc9c", "#e67e22"
    )
    ep_colours <- setNames(rep_len(base_colours, n_ep), ep_order)

    # Vertical jitter offsets so overlapping endpoints are legible
    offsets <- setNames(
      seq(-0.30, 0.30, length.out = max(n_ep, 1L)),
      ep_order
    )

    plot_height <- max(500L, length(cats) * 52L + 160L)
    fig <- plot_ly(height = plot_height, source = "forest")

    for (ep in ep_order) {
      dd <- dplyr::filter(d, endpoint_short == ep)
      if (nrow(dd) == 0L) next
      yy  <- unname(cat_y[dd$category]) + offsets[[ep]]
      col <- ep_colours[[ep]]

      pt_color <- ifelse(dd$interaction_p <= thresh, col,  "rgba(190,190,190,0.45)")
      bd_color <- ifelse(dd$interaction_p <= thresh, col,  "rgba(150,150,150,0.55)")

      hover_text <- paste0(
        "<b>", ep, "</b>",
        " <span style='color:#888;font-size:10px;'>[", dd$domain, "]</span><br>",
        "Subgroup category: <b>", dd$category, "</b><br>",
        "Values compared: ",      dd$subgroup_values, "<br>",
        "p(interaction) = <b>",  sprintf("%.4f", dd$interaction_p), "</b>",
        ifelse(
          dd$interaction_p <= thresh,
          paste0(" <span style='color:", col, ";'>\u25cf significant</span>"),
          ""
        ),
        "<br><i style='font-size:10px;color:#999;'>", dd$title, "</i>",
        "<br><span style='font-size:10px;color:#aaa;'>Double-click to load table in details panel</span>"
      )

      fig <- add_trace(fig,
        x           = dd$interaction_p,
        y           = yy,
        type        = "scatter",
        mode        = "markers",
        name        = ep,
        legendgroup = ep,
        # pass table_number as customdata for double-click identification
        customdata  = as.character(dd$table_number),
        marker      = list(
          size   = 12,
          color  = pt_color,
          symbol = "circle",
          line   = list(width = 1.5, color = bd_color)
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

    layout(fig,
      title  = list(
        text = "Interaction P-Values by Subgroup",
        font = list(size = 15, color = "#2c3e50"),
        x    = 0.02
      ),
      xaxis = list(
        title      = "Interaction P-Value",
        range      = c(-0.03, 1.05),
        zeroline   = FALSE,
        gridcolor  = "#ececec",
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
      # Legend placed vertically on the right – no longer overlaps X axis label
      legend = list(
        orientation = "v",
        x           = 1.02,
        y           = 1,
        xanchor     = "left",
        yanchor     = "top",
        font        = list(size = 11)
      ),
      hovermode     = "closest",
      plot_bgcolor  = "white",
      paper_bgcolor = "white",
      margin        = list(l = 260, r = 200, t = 60, b = 80)
    )
  })

  # ── Details table ─────────────────────────────────────────────────────────

  output$detail_table <- DT::renderDataTable({
    d      <- all_data()
    thresh <- input$threshold

    if (nrow(d) == 0L) {
      return(DT::datatable(data.frame(Message = "No data available."),
                           options = list(dom = "t"), rownames = FALSE))
    }

    if (input$detail_mode == "all") {
      # All rows from every source table that contains at least one significant p
      sig_tables <- d |>
        dplyr::filter(interaction_p <= thresh) |>
        dplyr::pull(table_number) |>
        unique()
      detail_d <- dplyr::filter(d, table_number %in% sig_tables)
    } else {
      # Rows from tables selected via double-click
      keys <- selected_tables()
      if (length(keys) == 0L) {
        return(DT::datatable(
          data.frame(Message = "Double-click any point on the plot to load its source table here."),
          options = list(dom = "t"), rownames = FALSE
        ))
      }
      sel_tn   <- as.integer(keys)
      detail_d <- dplyr::filter(d, table_number %in% sel_tn)
    }

    if (nrow(detail_d) == 0L) {
      return(DT::datatable(
        data.frame(Message = "No matching rows found."),
        options = list(dom = "t"), rownames = FALSE
      ))
    }

    # Build display table
    detail_d <- detail_d |>
      dplyr::arrange(table_number, interaction_p) |>
      dplyr::mutate(
        Significant   = ifelse(interaction_p <= thresh, "Yes", "No"),
        interaction_p = round(interaction_p, 4)
      ) |>
      dplyr::select(
        Domain             = domain,
        Endpoint           = endpoint_short,
        Category           = category,
        `Subgroup Values`  = subgroup_values,
        `p (interaction)`  = interaction_p,
        Significant,
        Table              = table_number,
        Title              = title
      )

    DT::datatable(
      detail_d,
      rownames  = FALSE,
      selection = "none",
      filter    = "top",
      options   = list(
        pageLength = 15,
        scrollX    = TRUE,
        dom        = "lftip",
        columnDefs = list(list(width = "35%", targets = 7))   # Title column wider
      )
    ) |>
      DT::formatStyle(
        "Significant",
        target          = "row",
        backgroundColor = DT::styleEqual(c("Yes", "No"), c("#fde8e8", "white")),
        fontWeight      = DT::styleEqual(c("Yes", "No"), c("bold", "normal"))
      )
  })
}

shinyApp(ui, server)
