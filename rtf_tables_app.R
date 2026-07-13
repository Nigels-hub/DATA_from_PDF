# ─────────────────────────────────────────────────────────────────────────────
# RTF Tabellen-Generator – Shiny App
#
# Liest extrahierte Excel-Tabellen aus extracted_tables/
# Erzeugt stilisierte RTF-Tabellen mit optionaler EN→DE-Übersetzung
#
# Requires: flextable >= 0.9.1, officer, readxl, DT, zip, shiny, dplyr, stringr
#
# Start:
#   Rscript -e "shiny::runApp('rtf_tables_app.R')"
# ─────────────────────────────────────────────────────────────────────────────

if (!requireNamespace("pacman", quietly = TRUE)) install.packages("pacman")
pacman::p_load(shiny, readxl, flextable, officer, dplyr, stringr, DT, tools, htmltools, zip)

# Optional: PDF page rendering for the PDF vs RTF comparison view.
# Requires the 'poppler' system library:  brew install poppler
# Then install R packages:               install.packages(c('pdftools', 'png'))
HAS_PDFTOOLS <- requireNamespace("pdftools", quietly = TRUE) &&
                requireNamespace("png",      quietly = TRUE)

EXCEL_DIR  <- "extracted_tables"
TRANS_FILE <- "translations_custom.csv"

# ══════════════════════════════════════════════════════════════════════════════
# 1 · TRANSLATION SYSTEM
# ══════════════════════════════════════════════════════════════════════════════

DEFAULT_TRANS <- data.frame(
  english = c(
    # Column headers
    "Parameter and Category/Statistic", "Parameter", "Category", "Statistic",
    # Statistics
    "n", "Mean", "SD", "Median", "Min", "Max",
    "95% CI", "95% Confidence Interval", "Hazard Ratio",
    "p-value", "P-value", "P-Value",
    "Overall", "Total", "Missing",
    "Not Estimable", "Not Available", "Not Applicable",
    "Yes", "No",
    # Demographics
    "Sex", "Male", "Female",
    "Age", "Age (years)", "Age Group",
    "Race", "Region",
    "Weight (kg)", "Height (cm)", "BMI (kg/m2)",
    # Efficacy
    "Overall Survival", "Progression-Free Survival",
    "Complete Response", "Partial Response",
    "Stable Disease", "Progressive Disease",
    "Overall Response", "Disease Control",
    "Events", "Censored", "Months", "Weeks", "Days",
    # Analysis sets
    "Full Analysis Set", "Safety Analysis Set", "Per Protocol Set",
    # Subjects
    "Subjects", "Patients", "Number of Subjects", "Number of Patients",
    # Table title components
    "Summary of", "Summary and Results of",
    "by Subgroups", "Completion Status",
    "Observed Means and Change from Baseline",
    "Time to First Deterioration",
    "Type of Events and Censoring",
    "Baseline Demographics", "Medical History",
    "Concomitant Therapies", "Study Drug Exposure",
    "Disposition of Subjects", "Duration of Observation",
    # AE terms
    "Adverse Event", "Serious Adverse Event",
    "Severe", "Fatal", "Deaths", "Discontinuation",
    "Treatment Discontinuation",
    "Subjects with at least one", "Subjects with no",
    # Other common terms
    "Subgroup", "Confidence Interval",
    "Standard Deviation", "Interquartile Range", "Range",
    "Number at risk", "Median survival time",
    "Lower", "Upper", "First", "Second", "Third",
    "Baseline", "Week", "Month", "Cycle",
    "Complete", "Partial", "None"
  ),
  german = c(
    "Parameter und Kategorie/Statistik", "Parameter", "Kategorie", "Statistik",
    "n", "Mittelwert", "SA", "Median", "Min", "Max",
    "95%-KI", "95%-Konfidenzintervall", "Hazard-Ratio",
    "p-Wert", "p-Wert", "p-Wert",
    "Gesamt", "Gesamt", "Fehlend",
    "Nicht schätzbar", "Nicht verfügbar", "Nicht anwendbar",
    "Ja", "Nein",
    "Geschlecht", "Männlich", "Weiblich",
    "Alter", "Alter (Jahre)", "Altersgruppe",
    "Ethnizität", "Region",
    "Körpergewicht (kg)", "Körpergröße (cm)", "BMI (kg/m²)",
    "Gesamtüberleben", "Progressionsfreies Überleben",
    "Vollständiges Ansprechen", "Partielles Ansprechen",
    "Stabile Erkrankung", "Progression",
    "Gesamtansprechen", "Krankheitskontrolle",
    "Ereignisse", "Zensiert", "Monate", "Wochen", "Tage",
    "Vollständiger Analysesatz (FAS)", "Sicherheits-Analysesatz",
    "Per-Protokoll-Analysesatz",
    "Probanden", "Patienten", "Anzahl Probanden", "Anzahl Patienten",
    "Zusammenfassung von", "Zusammenfassung und Ergebnisse von",
    "nach Untergruppen", "Vollständigkeitsstatus",
    "Beobachtete Mittelwerte und Veränderung gegenüber Ausgangswert",
    "Zeit bis zur ersten Verschlechterung",
    "Art der Ereignisse und Zensierung",
    "Baseline-Demografika", "Krankengeschichte",
    "Begleittherapien", "Studienmedikamenten-Exposition",
    "Probandenverteilung", "Beobachtungsdauer",
    "Unerwünschtes Ereignis (UE)", "Schwerwiegendes UE",
    "Schwer", "Fatal", "Todesfälle", "Abbruch",
    "Behandlungsabbruch",
    "Probanden mit mindestens einem", "Probanden ohne",
    "Untergruppe", "Konfidenzintervall",
    "Standardabweichung", "Interquartilsbereich", "Bereich",
    "Anzahl Patienten unter Risiko", "Medianes Überleben",
    "Untere", "Obere", "Erste", "Zweite", "Dritte",
    "Ausgangswert", "Woche", "Monat", "Zyklus",
    "Vollständig", "Partiell", "Keiner"
  ),
  stringsAsFactors = FALSE
)

load_trans <- function() {
  if (file.exists(TRANS_FILE)) {
    tryCatch({
      df <- read.csv(TRANS_FILE, stringsAsFactors = FALSE)
      if (all(c("english", "german") %in% names(df))) df else DEFAULT_TRANS
    }, error = function(e) DEFAULT_TRANS)
  } else {
    DEFAULT_TRANS
  }
}

save_trans <- function(df) {
  write.csv(df[, c("english", "german")], TRANS_FILE, row.names = FALSE)
}

make_trans_map <- function(df) setNames(df$german, df$english)

translate_text <- function(text, tmap) {
  if (is.null(text) || is.na(text) || !nzchar(trimws(text))) return(text)
  # 1. Exact match (fastest path)
  hit <- tmap[text]
  if (!is.na(hit)) return(unname(hit))
  # 2. Substring replacement – longer entries first (prevents partial overwrites)
  result <- text
  for (k in names(tmap)[order(nchar(names(tmap)), decreasing = TRUE)]) {
    if (grepl(k, result, fixed = TRUE))
      result <- gsub(k, tmap[[k]], result, fixed = TRUE)
  }
  result
}

translate_vec <- function(v, tmap) {
  vapply(v, translate_text, character(1), tmap = tmap, USE.NAMES = FALSE)
}

is_numeric_like <- function(v) {
  nzchar(v) &&
    grepl("^[\\d\\s\\.\\-\\+\\(\\)\\%\\<\\>\\,/NE]*$", v, perl = TRUE)
}

translate_df <- function(df, tmap) {
  result <- lapply(seq_len(ncol(df)), function(j) {
    vapply(df[[j]], function(v) {
      if (!nzchar(v) || is_numeric_like(v)) v else translate_text(v, tmap)
    }, character(1), USE.NAMES = FALSE)
  })
  out <- as.data.frame(result, stringsAsFactors = FALSE)
  colnames(out) <- colnames(df)
  out
}

# ══════════════════════════════════════════════════════════════════════════════
# 2 · EXCEL READER
# ══════════════════════════════════════════════════════════════════════════════

read_excel_table <- function(path) {
  raw <- suppressMessages(
    readxl::read_excel(path, col_names = FALSE, .name_repair = "minimal",
                       trim_ws = FALSE)
  )

  if (nrow(raw) == 0 || ncol(raw) == 0)
    return(list(title = basename(path), col_names = character(0),
                extra_hdrs = list(), data = data.frame(), footnotes = character(0)))

  to_s  <- function(x) { v <- as.character(x); ifelse(is.na(v) | v == "NA", "", v) }
  get_r <- function(i) vapply(seq_len(ncol(raw)), function(j) to_s(raw[[j]][i]), character(1))

  # Row 1 → title
  title <- to_s(raw[[1]][1])
  if (!nzchar(title)) title <- tools::file_path_sans_ext(basename(path))

  # Separator row (first col = "______...")
  sep_idx <- which(vapply(seq_len(nrow(raw)), function(i) {
    grepl("^_{5,}", to_s(raw[[1]][i]))
  }, logical(1)))
  data_end <- if (length(sep_idx)) sep_idx[1] - 1L else nrow(raw)

  # Footnotes
  footnotes <- character(0)
  if (length(sep_idx) && sep_idx[1] < nrow(raw)) {
    fn_rows <- seq(sep_idx[1] + 1L, nrow(raw))
    fn_vals <- vapply(fn_rows, function(i) to_s(raw[[1]][i]), character(1))
    footnotes <- fn_vals[nzchar(fn_vals)]
  }

  # Find main header row: scan rows 2..min(12, data_end) for "(N=" pattern
  header_end <- min(4L, data_end)   # safe default
  for (i in seq(2L, min(12L, data_end))) {
    row_text <- paste(get_r(i), collapse = " ")
    if (grepl("\\(N\\s*=", row_text)) { header_end <- i; break }
  }

  data_start <- header_end + 1L

  # Main column names (last header row)
  col_names_raw <- get_r(header_end)

  # Extra header rows between row 2 and header_end
  extra_hdrs <- list()
  for (i in seq(2L, header_end - 1L)) {
    rv <- get_r(i)
    if (any(nzchar(rv))) extra_hdrs[[length(extra_hdrs) + 1L]] <- rv
  }

  # Data block
  if (data_start > data_end || data_end < 1L) {
    df <- as.data.frame(
      matrix("", nrow = 0L, ncol = max(1L, length(col_names_raw))),
      stringsAsFactors = FALSE
    )
  } else {
    df_list <- lapply(seq_len(ncol(raw)), function(j) {
      vapply(seq(data_start, data_end), function(i) to_s(raw[[j]][i]), character(1))
    })
    df <- as.data.frame(df_list, stringsAsFactors = FALSE)
  }

  safe_cn <- make.unique(
    ifelse(nzchar(col_names_raw), col_names_raw, paste0("V", seq_along(col_names_raw)))
  )
  colnames(df) <- safe_cn

  # Remove fully empty rows
  if (nrow(df) > 0) {
    not_empty <- apply(df, 1, function(r) any(nzchar(r)))
    df <- df[not_empty, , drop = FALSE]
  }
  rownames(df) <- NULL

  list(title = title, col_names = col_names_raw,
       extra_hdrs = extra_hdrs, data = df, footnotes = footnotes)
}

# ══════════════════════════════════════════════════════════════════════════════
# 3 · TABLE STYLES  (4 styles)
# ══════════════════════════════════════════════════════════════════════════════
# Each style applies non-border properties; borders are applied separately
# so that the title header row (prepended last) does not interfere.

STYLES <- c(
  "Klinisch Standard"  = "clinical",
  "Kompakt"            = "compact",
  "Minimal (APA)"      = "apa",
  "Formales Raster"    = "grid"
)

# Apply font / colour / alignment / padding – NO borders
.style_props <- function(ft, style_id) {
  nc     <- length(ft$col_keys)
  data_j <- if (nc >= 2L) seq(2L, nc) else integer(0)
  nr_b   <- nrow_part(ft, "body")

  font_nm   <- if (style_id == "apa") "Times New Roman" else "Arial"
  font_sz   <- if (style_id == "compact") 8 else if (style_id == "apa") 10 else 9
  pad       <- if (style_id == "compact") 2 else 3
  pad_lr    <- pad + 1L

  ft <- font(ft,     fontname = font_nm, part = "all")
  ft <- fontsize(ft, size = font_sz,     part = "all")
  ft <- align(ft,    align = "left",     part = "all")
  ft <- padding(ft,  padding.top = pad, padding.bottom = pad,
                padding.left = pad_lr, padding.right = pad_lr, part = "all")
  ft <- border_remove(ft)

  if (length(data_j) > 0) {
    ft <- align(ft, j = data_j, align = "right",  part = "body")
    ft <- align(ft, j = data_j, align = "center", part = "header")
  }

  if (style_id == "clinical") {
    ft <- bold(ft,  bold = TRUE,     part = "header")
    ft <- color(ft, color = "white", part = "header")
    ft <- bg(ft,    bg = "#1a3a5c",  part = "header")

  } else if (style_id == "compact") {
    ft <- bold(ft,  bold = TRUE,     part = "header")
    ft <- color(ft, color = "white", part = "header")
    ft <- bg(ft,    bg = "#3d6da8",  part = "header")
    if (nr_b > 0L) {
      odd  <- seq(1L, nr_b, 2L)
      even <- seq(2L, nr_b, 2L)
      if (length(odd))  ft <- bg(ft, i = odd,  bg = "#f0f5fb", part = "body")
      if (length(even)) ft <- bg(ft, i = even, bg = "white",   part = "body")
    }

  } else if (style_id == "apa") {
    ft <- bold(ft,  bold = TRUE,   part = "header")
    ft <- color(ft, color = "black", part = "all")
    ft <- bg(ft,    bg = "white",    part = "all")

  } else { # grid
    ft <- bold(ft,  bold = TRUE,     part = "header")
    ft <- color(ft, color = "white", part = "header")
    ft <- bg(ft,    bg = "#1e4d8c",  part = "header")
    ft <- color(ft, color = "#111",  part = "body")
    ft <- bg(ft,    bg = "white",    part = "body")
  }

  ft <- set_table_properties(ft, layout = "autofit", width = 1)
  ft
}

# Apply ALL borders after the title header row has been prepended
.apply_borders <- function(ft, style_id) {
  b_thick <- fp_border(color = "#111111", width = 1.5)
  b_med   <- fp_border(color = "#444444", width = 1.0)
  b_thin  <- fp_border(color = "#aaaaaa", width = 0.5)
  b_none  <- fp_border(color = "transparent", width = 0)

  nr_h <- nrow_part(ft, "header")

  ft <- border_remove(ft)

  # ── Top of entire table ──
  ft <- hline_top(ft, border = b_thick, part = "header")

  # ── Separator between title row (i=1) and styled header rows ──
  if (nr_h > 1L) {
    ft <- border(ft, i = 1L,
                 border.bottom = fp_border(color = "#888888", width = 0.4),
                 part = "header")
  }

  # ── Bottom of header (separates header from body) ──
  ft <- hline_bottom(ft, border = b_thick, part = "header")

  # ── Bottom of table ──
  ft <- hline_bottom(ft, border = b_thick, part = "body")

  if (style_id == "compact") {
    ft <- hline(ft, border = b_thin, part = "body")
  }

  if (style_id == "grid") {
    ft <- border_outer(ft,   border = b_med,  part = "all")
    ft <- border_inner_h(ft, border = b_thin, part = "body")
    ft <- border_inner_v(ft, border = b_thin, part = "all")
  }

  ft
}

# ══════════════════════════════════════════════════════════════════════════════
# 4 · BUILD FLEXTABLE
# ══════════════════════════════════════════════════════════════════════════════

build_flextable <- function(tbl, style_id, lang = "EN", tmap = NULL) {

  # Apply translation
  if (lang == "DE" && !is.null(tmap)) {
    tbl$title      <- translate_text(tbl$title, tmap)
    tbl$col_names  <- translate_vec(tbl$col_names, tmap)
    tbl$extra_hdrs <- lapply(tbl$extra_hdrs, translate_vec, tmap = tmap)
    tbl$data       <- translate_df(tbl$data, tmap)
    tbl$footnotes  <- translate_vec(tbl$footnotes, tmap)
  }

  df <- tbl$data
  cn <- tbl$col_names

  # Fallback for empty tables
  if (nrow(df) == 0L || ncol(df) == 0L) {
    df <- data.frame(Hinweis = "Keine Datensätze gefunden.", check.names = FALSE)
    cn <- "Hinweis"
  }

  nc   <- length(cn)
  safe <- make.unique(ifelse(nzchar(cn), cn, paste0("V", seq_len(nc))))
  colnames(df) <- safe

  ft <- flextable(df)
  ft <- set_header_labels(ft, values = setNames(as.list(cn), safe))

  # ── Extra header rows (treatment group names etc.) ──────────────────────
  # add_header_row() prepends → use rev() so original order is preserved
  for (eh in rev(tbl$extra_hdrs)) {
    padded <- c(eh, rep("", max(0L, nc - length(eh))))[seq_len(nc)]
    ft <- add_header_row(ft, values = as.list(padded), colwidths = rep(1L, nc), top = TRUE)
    ft <- merge_h(ft, part = "header", i = 1L)  # merge identical adjacent cells
  }

  # ── Apply non-border style ───────────────────────────────────────────────
  ft <- .style_props(ft, style_id)

  # ── Apply indentation to first column based on leading spaces ─────────────
  # Detects "  Label" (2 spaces per indent level) and applies left padding
  # plus removes the leading spaces from the display text
  if (nrow(df) > 0L) {
    first_col <- df[[1L]]
    for (i in seq_len(nrow(df))) {
      txt <- first_col[i]
      if (nzchar(txt)) {
        n_spaces <- nchar(txt) - nchar(sub("^ +", "", txt))
        if (n_spaces > 0L) {
          indent_level <- as.integer(n_spaces / 2L)
          # Strip leading spaces from the cell
          cleaned_txt <- trimws(txt)
          # Apply left padding: add 3pt per indent level
          pad_left <- indent_level * 6L + 3L
          ft <- padding(ft, i = i, j = 1L, padding.left = pad_left, part = "body")
          # Update the cell value to show cleaned text
          ft <- compose(ft, i = i, j = 1L,
                        value = as_paragraph(cleaned_txt),
                        part = "body")
        }
      }
    }
  }

  # ── Prepend title row ────────────────────────────────────────────────────
  ft <- add_header_lines(ft, values = tbl$title)
  # i = 1 is now the title row → white background, black text, bold
  ft <- bg(ft,       i = 1L, bg = "white",         part = "header")
  ft <- color(ft,    i = 1L, color = "#111111",     part = "header")
  ft <- bold(ft,     i = 1L, bold = TRUE,           part = "header")
  ft <- italic(ft,   i = 1L, italic = FALSE,        part = "header")
  ft <- fontsize(ft, i = 1L, size = 9,              part = "header")
  ft <- align(ft,    i = 1L, align = "left",        part = "header")
  ft <- font(ft,     i = 1L, fontname = "Arial",    part = "header")

  # ── Apply borders (including the title row) ──────────────────────────────
  ft <- .apply_borders(ft, style_id)

  # ── Footnotes ────────────────────────────────────────────────────────────
  if (length(tbl$footnotes) > 0L) {
    ft <- add_footer_lines(ft, values = paste(tbl$footnotes, collapse = "\n"))
    ft <- italic(ft,   part = "footer")
    ft <- fontsize(ft, size = 7.5,       part = "footer")
    ft <- color(ft,    color = "#555555", part = "footer")
    ft <- bg(ft,       bg = "white",      part = "footer")
    ft <- border(ft,
                 border.top    = fp_border(color = "#888888", width = 0.4),
                 border.bottom = fp_border(color = "transparent", width = 0),
                 part = "footer")
  }

  ft
}

# ══════════════════════════════════════════════════════════════════════════════
# 5 · FILE LIST
# ══════════════════════════════════════════════════════════════════════════════

get_file_list <- function(dir = EXCEL_DIR) {
  fs <- list.files(dir, pattern = "\\.xlsx$", full.names = FALSE)
  if (length(fs) == 0L)
    return(data.frame(Nr = integer(0), Tabellen_ID = character(0),
                      Titel = character(0), Datei = character(0),
                      stringsAsFactors = FALSE))

  nr   <- suppressWarnings(as.integer(sub("^(\\d+)_.*$", "\\1", fs)))
  rest <- sub("^\\d+_", "", sub("\\.xlsx$", "", fs))
  tid  <- sub("_.*$", "", rest)
  ttl  <- sub("^[^_]+_\\s*", "", rest)

  data.frame(Nr = nr, Tabellen_ID = tid, Titel = ttl, Datei = fs,
             stringsAsFactors = FALSE)
}

# ══════════════════════════════════════════════════════════════════════════════
# 6 · UI
# ══════════════════════════════════════════════════════════════════════════════

ui <- fluidPage(
  tags$head(tags$style(HTML("
    body          { font-family:'Segoe UI',Arial,sans-serif; background:#f4f6f9; }
    .well         { background:white; border:1px solid #dee2e6; border-radius:6px; }
    .section-lbl  { font-weight:600; color:#2c3e50; font-size:13px;
                    margin-bottom:6px; margin-top:2px; }
    .preview-box  { background:white; border-radius:6px; padding:16px;
                    box-shadow:0 1px 4px rgba(0,0,0,.1); overflow-x:auto; }
    .dl-btn       { width:100%; margin-top:6px; }
    .info-box     { background:#f8f9fa; border-radius:4px; padding:8px 10px;
                    font-size:11px; color:#555; line-height:1.6; }
    table.dataTable { font-size:12px; }
    .tab-content  { padding-top:10px; }
    /* PDF vs RTF comparison modal */
    .modal-xl .modal-dialog { max-width:1380px; width:95vw; }
    .cmp-panel  { overflow-y:auto; max-height:74vh; }
    .cmp-no-pdf { color:#c0392b; font-size:12px; padding:20px;
                  text-align:center; background:#fff5f5; border-radius:4px; }
  "))),

  titlePanel(
    div(style = "color:#2c3e50;",
        "RTF Tabellen-Generator",
        tags$small(style = "font-size:13px; color:#7f8c8d; margin-left:12px;",
                   "Zolbetuximab + mFOLFOX6 | Klinische Studienergebnisse"))
  ),

  sidebarLayout(

    # ── Sidebar ──────────────────────────────────────────────────────────────
    sidebarPanel(
      width = 3,

      div(class = "section-lbl", "Tabellenstil"),
      selectInput("style_id", label = NULL,
                  choices = STYLES, selected = "clinical"),

      div(class = "section-lbl", "Sprache / Language"),
      radioButtons("language", label = NULL,
                   choices  = c("Englisch (EN)" = "EN",
                                "Deutsch (DE)"  = "DE"),
                   selected = "EN", inline = TRUE),

      hr(),

      div(class = "section-lbl", "Download"),
      downloadButton("dl_single", "Aktive Tabelle (RTF)",
                     class = "btn-primary dl-btn"),
      tags$br(),
      helpText(style = "font-size:11px; color:#888; margin-top:4px;",
               "Für Batch-Export → Tab \"Batch-Export\""),

      actionButton("btn_compare", "PDF vs RTF",
                   icon  = icon("table-columns"),
                   class = "btn-info dl-btn",
                   style = "margin-top:6px;"),
      helpText(style = "font-size:11px; color:#888; margin-top:4px;",
               "Original-PDF ↔ RTF-Vorschau nebeneinander"),

      hr(),

      div(class = "section-lbl", "Tabelleninfo"),
      div(class = "info-box", verbatimTextOutput("tbl_info"))
    ),

    # ── Main panel ───────────────────────────────────────────────────────────
    mainPanel(
      width = 9,
      tabsetPanel(
        id = "main_tabs",

        # ── Tab 1: Browser + Vorschau ──────────────────────────────────────
        tabPanel(
          "Tabellen & Vorschau",
          br(),
          div(class = "section-lbl",
              "Tabelle auswählen (Suchfeld oben in der Tabelle nutzbar):"),
          DTOutput("file_tbl"),
          hr(),
          div(class = "section-lbl", "HTML-Vorschau der ausgewählten Tabelle:"),
          div(class = "preview-box", uiOutput("ft_preview"))
        ),

        # ── Tab 2: Übersetzungen ───────────────────────────────────────────
        tabPanel(
          "Übersetzungen",
          br(),
          fluidRow(
            column(12,
              div(style = "margin-bottom:10px;",
                actionButton("btn_save_trans",  "Speichern",
                             icon = icon("floppy-disk"),
                             class = "btn-success btn-sm"),
                actionButton("btn_reset_trans", "Zurücksetzen",
                             icon = icon("rotate-left"),
                             class = "btn-warning btn-sm",
                             style = "margin-left:6px;"),
                actionButton("btn_add_row",     "Zeile hinzufügen",
                             icon = icon("plus"),
                             class = "btn-info btn-sm",
                             style = "margin-left:6px;"),
                actionButton("btn_del_rows",    "Auswahl löschen",
                             icon = icon("trash"),
                             class = "btn-danger btn-sm",
                             style = "margin-left:6px;")
              ),
              helpText(style = "font-size:11px; color:#666;",
                       "Doppelklick auf eine Zelle zum Bearbeiten. ",
                       "Änderungen wirken sofort auf Vorschau und Download. ",
                       "\"Speichern\" schreibt in translations_custom.csv."),
              DTOutput("trans_tbl")
            )
          )
        ),

        # ── Tab 3: Batch-Export ────────────────────────────────────────────
        tabPanel(
          "Batch-Export",
          br(),
          fluidRow(
            column(4,
              wellPanel(
                div(class = "section-lbl", "Batch-Optionen"),
                selectInput("b_style", "Stil", choices = STYLES, selected = "clinical"),
                radioButtons("b_lang", "Sprache",
                             choices = c("EN" = "EN", "DE" = "DE"),
                             selected = "EN", inline = TRUE),
                hr(),
                div(style = "margin-bottom:6px;",
                  actionButton("batch_all",  "Alle wählen",
                               class = "btn-sm btn-default"),
                  actionButton("batch_none", "Alle abwählen",
                               class = "btn-sm btn-default",
                               style = "margin-left:4px;")
                ),
                uiOutput("batch_sel_info"),
                hr(),
                downloadButton("dl_batch", "ZIP herunterladen",
                               class = "btn-primary",
                               style = "width:100%;"),
                helpText(style = "font-size:11px; color:#888; margin-top:4px;",
                         "Erzeugt eine ZIP-Datei mit allen ausgewählten RTF-Tabellen.")
              )
            ),
            column(8,
              div(class = "section-lbl", "Tabellen für Batch-Export auswählen:"),
              DTOutput("batch_tbl")
            )
          )
        )
      )
    )
  )
)

# ══════════════════════════════════════════════════════════════════════════════
# 7 · SERVER
# ══════════════════════════════════════════════════════════════════════════════

server <- function(input, output, session) {

  # ── File list ─────────────────────────────────────────────────────────────
  file_df <- reactive({ get_file_list() })

  output$file_tbl <- renderDT({
    df <- file_df()[, c("Nr", "Tabellen_ID", "Titel")]
    datatable(
      df,
      selection = "single",
      rownames  = FALSE,
      filter    = "top",
      options   = list(
        pageLength = 10,
        scrollX    = TRUE,
        columnDefs = list(
          list(width = "50px",  targets = 0L),
          list(width = "120px", targets = 1L)
        )
      ),
      class = "compact stripe hover"
    )
  })

  selected_path <- reactive({
    rows <- input$file_tbl_rows_selected
    if (is.null(rows) || length(rows) == 0L) return(NULL)
    file.path(EXCEL_DIR, file_df()$Datei[rows])
  })

  # ── Translation ───────────────────────────────────────────────────────────
  trans_rv <- reactiveVal(load_trans())

  output$trans_tbl <- renderDT({
    datatable(
      trans_rv(),
      editable  = "cell",
      rownames  = FALSE,
      selection = "multiple",
      options   = list(pageLength = 20, scrollX = TRUE),
      colnames  = c("Englisch (Original)", "Deutsch (Übersetzung)"),
      class     = "compact stripe"
    )
  })

  observeEvent(input$trans_tbl_cell_edit, {
    info <- input$trans_tbl_cell_edit
    df   <- trans_rv()
    df[info$row, info$col + 1L] <- info$value
    trans_rv(df)
  })

  observeEvent(input$btn_save_trans, {
    save_trans(trans_rv())
    showNotification("Übersetzungen in translations_custom.csv gespeichert.",
                     type = "message", duration = 3)
  })

  observeEvent(input$btn_reset_trans, {
    trans_rv(DEFAULT_TRANS)
    showNotification("Standardübersetzungen wiederhergestellt.",
                     type = "warning", duration = 3)
  })

  observeEvent(input$btn_add_row, {
    df <- trans_rv()
    df <- rbind(df, data.frame(english = "New Term", german = "Neuer Begriff",
                               stringsAsFactors = FALSE))
    trans_rv(df)
  })

  observeEvent(input$btn_del_rows, {
    sel <- input$trans_tbl_rows_selected
    if (!is.null(sel) && length(sel) > 0L) {
      df <- trans_rv()
      trans_rv(df[-sel, , drop = FALSE])
    }
  })

  tmap_rv <- reactive({ make_trans_map(trans_rv()) })

  # ── Current table (read on demand) ───────────────────────────────────────
  current_tbl <- reactive({
    fp <- selected_path()
    req(fp)
    withProgress(message = "Lese Tabelle …", value = 0.5,
                 read_excel_table(fp))
  })

  # ── Table info ────────────────────────────────────────────────────────────
  output$tbl_info <- renderText({
    fp <- selected_path()
    if (is.null(fp)) return("Keine Tabelle ausgewählt.")
    tbl <- current_tbl()
    sprintf("Zeilen:     %d\nSpalten:    %d\nFußnoten:   %d\nExtra-Hdr:  %d",
            nrow(tbl$data), ncol(tbl$data),
            length(tbl$footnotes), length(tbl$extra_hdrs))
  })

  # ── Preview ───────────────────────────────────────────────────────────────
  output$ft_preview <- renderUI({
    fp <- selected_path()
    if (is.null(fp)) {
      return(div(style = "color:#aaa; padding:30px; text-align:center;",
                 icon("table"), " Bitte eine Tabelle aus der Liste oben auswählen."))
    }
    tbl <- current_tbl()
    ft  <- tryCatch(
      build_flextable(tbl, input$style_id, input$language, tmap_rv()),
      error = function(e) {
        return(div(class = "alert alert-danger",
                   paste("Fehler beim Aufbau der Tabelle:", e$message)))
      }
    )
    if (inherits(ft, "shiny.tag")) return(ft)
    htmltools_value(ft)
  })

  # ── Single RTF download ───────────────────────────────────────────────────
  output$dl_single <- downloadHandler(
    filename = function() {
      fp <- selected_path()
      if (is.null(fp)) return("tabelle.rtf")
      paste0(tools::file_path_sans_ext(basename(fp)),
             "_", input$style_id, "_", input$language, ".rtf")
    },
    content = function(file) {
      tbl <- current_tbl()
      ft  <- build_flextable(tbl, input$style_id, input$language, tmap_rv())
      save_as_rtf(ft, path = file)
    }
  )

  # ── Batch export ──────────────────────────────────────────────────────────
  output$batch_tbl <- renderDT({
    df <- file_df()[, c("Nr", "Tabellen_ID", "Titel")]
    datatable(
      df,
      selection = "multiple",
      rownames  = FALSE,
      filter    = "top",
      options   = list(pageLength = 15, scrollX = TRUE,
                       columnDefs = list(list(width = "50px",  targets = 0L),
                                         list(width = "120px", targets = 1L))),
      class = "compact stripe hover"
    )
  })

  observeEvent(input$batch_all, {
    proxy <- dataTableProxy("batch_tbl")
    selectRows(proxy, seq_len(nrow(file_df())))
  })

  observeEvent(input$batch_none, {
    proxy <- dataTableProxy("batch_tbl")
    selectRows(proxy, NULL)
  })

  output$batch_sel_info <- renderUI({
    n <- length(input$batch_tbl_rows_selected)
    div(class = "info-box",
        sprintf("%d Tabelle(n) ausgewählt", n))
  })

  output$dl_batch <- downloadHandler(
    filename = function() {
      paste0("RTF_Tabellen_", input$b_style, "_", input$b_lang, "_",
             format(Sys.Date(), "%Y%m%d"), ".zip")
    },
    content = function(file) {
      rows <- input$batch_tbl_rows_selected
      if (is.null(rows) || length(rows) == 0L) {
        showNotification("Keine Tabellen ausgewählt.", type = "warning")
        writeLines("No tables selected.", file)
        return()
      }

      df   <- file_df()
      tmap <- tmap_rv()
      bst  <- input$b_style
      bln  <- input$b_lang

      tmp_dir <- tempfile(pattern = "rtf_batch_")
      dir.create(tmp_dir, recursive = TRUE)

      n_total <- length(rows)
      errors  <- character(0)

      withProgress(message = "Erzeuge RTF-Dateien …", value = 0, {
        for (idx in seq_along(rows)) {
          r  <- rows[idx]
          fp <- file.path(EXCEL_DIR, df$Datei[r])
          tryCatch({
            tbl <- read_excel_table(fp)
            ft  <- build_flextable(tbl, bst, bln, tmap)
            out <- file.path(tmp_dir,
                             paste0(tools::file_path_sans_ext(df$Datei[r]),
                                    "_", bst, ".rtf"))
            save_as_rtf(ft, path = out)
          }, error = function(e) {
            errors <<- c(errors, sprintf("%s: %s", df$Datei[r], e$message))
          })
          incProgress(1 / n_total, detail = sprintf("%d / %d", idx, n_total))
        }
      })

      rtf_files <- list.files(tmp_dir, pattern = "\\.rtf$", full.names = FALSE)
      if (length(rtf_files) == 0L) {
        writeLines(c("No RTF files generated.", errors), file)
        return()
      }

      old_wd <- getwd()
      setwd(tmp_dir)
      tryCatch(
        utils::zip(zipfile = file, files = rtf_files),
        finally = setwd(old_wd)
      )

      if (length(errors) > 0L)
        showNotification(
          sprintf("%d Fehler aufgetreten. Details in der Konsole.", length(errors)),
          type = "warning", duration = 8
        )
    }
  )

  # ══════════════════════════════════════════════════════════════════════════
  # PDF vs RTF comparison
  # ══════════════════════════════════════════════════════════════════════════

  compare_state <- reactiveValues(page_num = NA_integer_)

  # Page index written by extract_tables.py alongside the Excel files
  page_index_rv <- reactive({
    idx_path <- file.path(EXCEL_DIR, "page_index.csv")
    if (!file.exists(idx_path)) return(NULL)
    tryCatch(read.csv(idx_path, stringsAsFactors = FALSE), error = function(e) NULL)
  })

  # Auto-detect the PDF in the working directory (largest file = main PDF)
  get_pdf_path <- function() {
    pdfs <- list.files(".", pattern = "\\.pdf$", full.names = TRUE)
    if (length(pdfs) == 0L) return(NULL)
    pdfs[which.max(file.info(pdfs)$size)]
  }

  observeEvent(input$btn_compare, {
    fp <- selected_path()
    if (is.null(fp)) {
      showNotification("Bitte zuerst eine Tabelle auswählen.",
                       type = "warning", duration = 3)
      return()
    }
    if (!HAS_PDFTOOLS) {
      showNotification(
        paste0("pdftools nicht verfügbar. Aktivieren mit: ",
               "brew install poppler && ",
               "Rscript -e \"install.packages(c('pdftools','png'))\""),
        type = "error", duration = 12)
      return()
    }

    pidx    <- isolate(page_index_rv())
    fname   <- basename(fp)
    has_idx <- !is.null(pidx) && fname %in% pidx$filename
    page_num <- if (has_idx)
                  as.integer(pidx$page_start[pidx$filename == fname][1L])
                else NA_integer_

    compare_state$page_num <- if (!is.na(page_num)) page_num else 1L

    showModal(modalDialog(
      title = tagList(
        icon("table-columns"), "\u2002PDF vs RTF\u2002\u2014\u2002",
        tags$span(style = "font-size:12px; font-weight:400; color:#999;", fname)
      ),
      size      = "xl",
      easyClose = TRUE,
      footer    = modalButton("Schlie\u00dfen"),

      fluidRow(
        column(6,
          div(class = "section-lbl", "Original PDF"),
          if (has_idx) {
            div(style = "font-size:11px; color:#777; margin-bottom:6px;",
                sprintf("Seite\u00a0%d", page_num))
          } else {
            div(style = "display:flex; align-items:center; gap:6px; margin-bottom:8px;",
              div(class = "alert alert-info",
                  style = "font-size:11px; padding:5px 10px; margin:0; flex:1;",
                  icon("circle-info"),
                  " Kein Seitenindex \u2014 bitte Seite manuell eingeben:"),
              numericInput("cmp_page_in", NULL, value = 1L,
                           min = 1L, step = 1L, width = "80px"),
              actionButton("cmp_go", "Zeigen",
                           class = "btn-sm btn-primary",
                           style = "white-space:nowrap;")
            )
          },
          div(class = "cmp-panel",
              imageOutput("cmp_pdf_img", height = "auto", width = "100%"))
        ),
        column(6,
          div(class = "section-lbl", "RTF-Vorschau"),
          div(class = "preview-box cmp-panel",
              uiOutput("cmp_rtf_preview"))
        )
      )
    ))
  })

  observeEvent(input$cmp_go, {
    req(input$cmp_page_in)
    compare_state$page_num <- as.integer(input$cmp_page_in)
  })

  output$cmp_rtf_preview <- renderUI({
    fp <- selected_path()
    if (is.null(fp)) return(NULL)
    tbl <- current_tbl()
    ft  <- tryCatch(
      build_flextable(tbl, input$style_id, input$language, tmap_rv()),
      error = function(e) NULL
    )
    if (is.null(ft))
      return(div(class = "alert alert-danger", "Fehler beim Aufbau der Tabelle."))
    htmltools_value(ft)
  })

  output$cmp_pdf_img <- renderImage({
    pn <- compare_state$page_num
    req(!is.na(pn) && pn >= 1L)

    tmp    <- tempfile(fileext = ".png")
    pdf_fp <- get_pdf_path()

    if (is.null(pdf_fp)) {
      grDevices::png(tmp, width = 500, height = 200, bg = "#fff5f5")
      par(mar = rep(0, 4)); plot.new()
      text(0.5, 0.5, "PDF-Datei nicht gefunden.",
           col = "#c0392b", cex = 1.3, font = 2)
      grDevices::dev.off()
    } else {
      img <- tryCatch(
        pdftools::pdf_render_page(pdf_fp, page = pn, dpi = 130L),
        error = function(e) NULL
      )
      if (!is.null(img)) {
        png::writePNG(img, target = tmp)
      } else {
        grDevices::png(tmp, width = 500, height = 200, bg = "#fff5f5")
        par(mar = rep(0, 4)); plot.new()
        text(0.5, 0.5, sprintf("Seite\u00a0%d nicht renderbar.", pn),
             col = "#c0392b", cex = 1.3, font = 2)
        grDevices::dev.off()
      }
    }

    list(src         = tmp,
         contentType = "image/png",
         alt         = paste("PDF-Seite", pn),
         style       = "max-width:100%; height:auto; border:1px solid #ddd;")
  }, deleteFile = TRUE)
}

# ══════════════════════════════════════════════════════════════════════════════
# 8 · LAUNCH
# ══════════════════════════════════════════════════════════════════════════════

shinyApp(ui, server)
