# ─────────────────────────────────────────────────────────────────────────────
# RTF Tabellen-Generator – Shiny App
#
# Liest extrahierte Excel-Tabellen aus extracted_tables/
# Erzeugt stilisierte RTF-Tabellen mit optionaler EN→DE-Übersetzung
# Speichert Tabellenstile und Konfigurationen als Muster
# Speichert Tabellenstile und Konfigurationen als Muster
#
# Requires: flextable >= 0.9.1, officer, readxl, DT, zip, shiny, dplyr, stringr, colourpicker
# Requires: flextable >= 0.9.1, officer, readxl, DT, zip, shiny, dplyr, stringr, colourpicker
#
# Start:
#   Rscript -e "shiny::runApp('rtf_tables_app.R')"
# ─────────────────────────────────────────────────────────────────────────────

# ── Dependency Management ────────────────────────────────────────────────────

# Install pacman for package management
if (!requireNamespace("pacman", quietly = TRUE)) {
  install.packages("pacman", repos = "https://cloud.r-project.org/")
}

# Define required R packages
REQUIRED_R_PACKAGES <- c(
  "shiny", "readxl", "flextable", "officer", "dplyr", "stringr",
  "DT", "tools", "htmltools", "zip", "colourpicker"
)

# Function to check and install R packages
check_and_install_r_packages <- function(packages = REQUIRED_R_PACKAGES) {
  missing <- packages[!sapply(packages, requireNamespace, quietly = TRUE)]
  if (length(missing) > 0L) {
    message("Installing missing R packages: ", paste(missing, collapse = ", "))
    pacman::p_load(char = missing)
  }
  invisible(TRUE)
}

# Function to check Python and required Python packages
check_python_environment <- function() {
  # Check if Python is available
  python_check <- tryCatch({
    result <- suppressWarnings(system2("python3", "--version", stdout = TRUE, stderr = TRUE))
    list(available = TRUE, version = paste(result, collapse = " "))
  }, error = function(e) {
    list(available = FALSE, version = NA_character_)
  })
  
  if (!python_check$available) {
    python_check <- tryCatch({
      result <- suppressWarnings(system2("python", "--version", stdout = TRUE, stderr = TRUE))
      list(available = TRUE, version = paste(result, collapse = " "))
    }, error = function(e) {
      list(available = FALSE, version = NA_character_)
    })
  }
  
  if (!python_check$available) {
    return(list(
      python_available = FALSE,
      packages = list(),
      error = "Python 3 not found. Please install Python 3."
    ))
  }
  
  # Check required Python packages by checking exit code only
  required_packages <- c("pdfplumber", "openpyxl")
  missing_packages <- character(0)
  
  for (pkg in required_packages) {
    # Simple approach: try to import and check exit code
    exit_code <- tryCatch({
      suppressWarnings(system2("python3", c("-c", paste0("import ", pkg)), 
                              stdout = FALSE, stderr = FALSE))
    }, error = function(e) {
      1  # Return error code if tryCatch fails
    })
    
    if (exit_code != 0) {
      missing_packages <- c(missing_packages, pkg)
    }
  }
  
  list(
    python_available = TRUE,
    python_version = python_check$version,
    required_packages = required_packages,
    missing_packages = missing_packages,
    all_packages_available = length(missing_packages) == 0L
  )
}

# Function to install missing Python packages
install_python_packages <- function(packages) {
  if (length(packages) == 0L) return(TRUE)
  
  for (pkg in packages) {
    message("Installing Python package: ", pkg)
    result <- system2("pip3", c("install", pkg), stdout = TRUE, stderr = TRUE)
    if (!any(grepl("Successfully installed", result))) {
      # Try with pip as fallback
      result <- system2("pip", c("install", pkg), stdout = TRUE, stderr = TRUE)
    }
  }
  invisible(TRUE)
}

# Run checks on startup
check_and_install_r_packages()
python_env_status <- check_python_environment()

# Optional: PDF page rendering for the PDF vs RTF comparison view.
# Requires the 'poppler' system library:  brew install poppler
# Then install R packages:               install.packages(c('pdftools', 'png'))
HAS_PDFTOOLS <- requireNamespace("pdftools", quietly = TRUE) &&
                requireNamespace("png",      quietly = TRUE)

# Load all required packages
pacman::p_load(shiny, readxl, flextable, officer, dplyr, stringr, DT, tools, htmltools, zip, colourpicker)

# Load all required packages
pacman::p_load(shiny, readxl, flextable, officer, dplyr, stringr, DT, tools, htmltools, zip, colourpicker)

EXCEL_DIR  <- "extracted_tables"
TRANS_FILE <- "translations_custom.csv"
FMT_RULES_FILE <- "formatting_rules_de.csv"
PATTERNS_DIR <- "muster"  # Directory for saving pattern templates

# ══════════════════════════════════════════════════════════════════════════════
# 1 · TRANSLATION SYSTEM
# ══════════════════════════════════════════════════════════════════════════════

# ── Formatting Rules for German (DE) ──────────────────────────────────────────
# Each rule applies a regex pattern substitution to numeric and percentage values
# when translating to German. Format: list(pattern = "regex_pattern", replacement = "replacement_string")
#
# Examples:
#   Decimal separator: pattern = "\\.", replacement = ","
#   Percentage spacing: pattern = "(%)(\\s*)", replacement = " %"
#
FORMATTING_RULES_DE <- list(
  list(
    name = "Decimal separator (. to ,)",
    pattern = "\\.",
    replacement = ",",
    description = "Converts English decimal separator (.) to German (,)",
    enabled = TRUE
  ),
  list(
    name = "Percentage spacing (add space before %)",
    pattern = "(\\d)\\s*(%)",
    replacement = "\\1 \\2",
    description = "Adds fixed space between number and % symbol",
    enabled = TRUE
  )
  # Add more rules here as needed, e.g.:
  # list(
  #   name = "Custom rule name",
  #   pattern = "your_regex_pattern",
  #   replacement = "your_replacement",
  #   description = "Description of what this rule does",
  #   enabled = TRUE
  # )
)

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

# ── Formatting rules (list ↔ dataframe conversion) ──────────────────────────
rules_to_df <- function(rules) {
  if (is.null(rules) || length(rules) == 0L) {
    return(data.frame(
      name = character(0),
      pattern = character(0),
      replacement = character(0),
      description = character(0),
      enabled = logical(0),
      stringsAsFactors = FALSE
    ))
  }
  df_list <- list(
    name = character(length(rules)),
    pattern = character(length(rules)),
    replacement = character(length(rules)),
    description = character(length(rules)),
    enabled = logical(length(rules))
  )
  for (i in seq_along(rules)) {
    df_list$name[i] <- if (is.null(rules[[i]]$name)) "" else rules[[i]]$name
    df_list$pattern[i] <- if (is.null(rules[[i]]$pattern)) "" else rules[[i]]$pattern
    df_list$replacement[i] <- if (is.null(rules[[i]]$replacement)) "" else rules[[i]]$replacement
    df_list$description[i] <- if (is.null(rules[[i]]$description)) "" else rules[[i]]$description
    df_list$enabled[i] <- isTRUE(rules[[i]]$enabled)
  }
  data.frame(df_list, stringsAsFactors = FALSE)
}

df_to_rules <- function(df) {
  if (nrow(df) == 0L) return(list())
  rules <- list()
  for (i in seq_len(nrow(df))) {
    rules[[i]] <- list(
      name = df$name[i],
      pattern = df$pattern[i],
      replacement = df$replacement[i],
      description = df$description[i],
      enabled = isTRUE(df$enabled[i])
    )
  }
  rules
}

load_fmt_rules <- function() {
  if (file.exists(FMT_RULES_FILE)) {
    tryCatch({
      df <- read.csv(FMT_RULES_FILE, stringsAsFactors = FALSE)
      if (all(c("name", "pattern", "replacement", "description", "enabled") %in% names(df))) {
        df$enabled <- as.logical(df$enabled)
        df_to_rules(df)
      } else {
        FORMATTING_RULES_DE
      }
    }, error = function(e) FORMATTING_RULES_DE)
  } else {
    FORMATTING_RULES_DE
  }
}

save_fmt_rules <- function(df) {
  write.csv(df[, c("name", "pattern", "replacement", "description", "enabled")],
            FMT_RULES_FILE, row.names = FALSE)
}

# ── Pattern (Muster) Management ──────────────────────────────────────────────
# Save/load complete pattern templates including translations, rules, and style
init_patterns_dir <- function() {
  if (!dir.exists(PATTERNS_DIR)) {
    dir.create(PATTERNS_DIR, recursive = TRUE)
  }
}

save_pattern <- function(pattern_name, trans_df, fmt_rules_df, style_id, custom_styles_df = NULL) {
  init_patterns_dir()
  pattern_path <- file.path(PATTERNS_DIR, pattern_name)
  
  if (!dir.exists(pattern_path)) {
    dir.create(pattern_path, recursive = TRUE)
  }
  
  # Save metadata (style)
  write.csv(data.frame(style = style_id, created = format(Sys.time())),
            file.path(pattern_path, "pattern_info.csv"), row.names = FALSE)
  
  # Save translations
  write.csv(trans_df[, c("english", "german")],
            file.path(pattern_path, "translations.csv"), row.names = FALSE)
  
  # Save formatting rules
  write.csv(fmt_rules_df[, c("name", "pattern", "replacement", "description", "enabled")],
            file.path(pattern_path, "formatting_rules.csv"), row.names = FALSE)
  
  # Save custom styles if provided
  if (!is.null(custom_styles_df) && nrow(custom_styles_df) > 0L) {
    write.csv(custom_styles_df, file.path(pattern_path, "custom_styles.csv"), row.names = FALSE)
  }
}

load_pattern <- function(pattern_name) {
  pattern_path <- file.path(PATTERNS_DIR, pattern_name)
  
  if (!dir.exists(pattern_path)) {
    return(NULL)
  }
  
  pattern_data <- list()
  
  # Load metadata
  tryCatch({
    info <- read.csv(file.path(pattern_path, "pattern_info.csv"), stringsAsFactors = FALSE)
    pattern_data$style <- if (nrow(info) > 0) info$style[1] else "clinical"
  }, error = function(e) {
    pattern_data$style <<- "clinical"
  })
  
  # Load translations
  tryCatch({
    trans <- read.csv(file.path(pattern_path, "translations.csv"), stringsAsFactors = FALSE)
    if (all(c("english", "german") %in% names(trans))) {
      pattern_data$translations <- trans
    } else {
      pattern_data$translations <- NULL
    }
  }, error = function(e) {
    pattern_data$translations <<- NULL
  })
  
  # Load formatting rules
  tryCatch({
    fmt <- read.csv(file.path(pattern_path, "formatting_rules.csv"), stringsAsFactors = FALSE)
    if (all(c("name", "pattern", "replacement", "description", "enabled") %in% names(fmt))) {
      fmt$enabled <- as.logical(fmt$enabled)
      pattern_data$fmt_rules <- fmt
    } else {
      pattern_data$fmt_rules <- NULL
    }
  }, error = function(e) {
    pattern_data$fmt_rules <<- NULL
  })
  
  # Load custom styles
  tryCatch({
    custom_styles <- read.csv(file.path(pattern_path, "custom_styles.csv"), stringsAsFactors = FALSE)
    # Convert to proper types
    if ("header_bold" %in% names(custom_styles)) {
      custom_styles$header_bold <- as.logical(custom_styles$header_bold)
    }
    if ("alternating_rows" %in% names(custom_styles)) {
      custom_styles$alternating_rows <- as.logical(custom_styles$alternating_rows)
    }
    if ("font_size" %in% names(custom_styles)) {
      custom_styles$font_size <- as.numeric(custom_styles$font_size)
    }
    if ("padding" %in% names(custom_styles)) {
      custom_styles$padding <- as.numeric(custom_styles$padding)
    }
    pattern_data$custom_styles <- custom_styles
  }, error = function(e) {
    pattern_data$custom_styles <<- NULL
  })
  
  # Load custom styles
  tryCatch({
    custom_styles <- read.csv(file.path(pattern_path, "custom_styles.csv"), stringsAsFactors = FALSE)
    # Convert to proper types
    if ("header_bold" %in% names(custom_styles)) {
      custom_styles$header_bold <- as.logical(custom_styles$header_bold)
    }
    if ("alternating_rows" %in% names(custom_styles)) {
      custom_styles$alternating_rows <- as.logical(custom_styles$alternating_rows)
    }
    if ("font_size" %in% names(custom_styles)) {
      custom_styles$font_size <- as.numeric(custom_styles$font_size)
    }
    if ("padding" %in% names(custom_styles)) {
      custom_styles$padding <- as.numeric(custom_styles$padding)
    }
    pattern_data$custom_styles <- custom_styles
  }, error = function(e) {
    pattern_data$custom_styles <<- NULL
  })
  
  pattern_data
}

delete_pattern <- function(pattern_name) {
  pattern_path <- file.path(PATTERNS_DIR, pattern_name)
  if (dir.exists(pattern_path)) {
    unlink(pattern_path, recursive = TRUE)
    return(TRUE)
  }
  FALSE
}

list_patterns <- function() {
  init_patterns_dir()
  patterns <- list.dirs(PATTERNS_DIR, full.names = FALSE, recursive = FALSE)
  patterns[patterns != ""]  # Remove empty strings
}

# ── Custom Styles Management ─────────────────────────────────────────────────
# Store custom table styles that users can define and save

# Default empty custom styles dataframe
get_empty_custom_styles_df <- function() {
  data.frame(
    name = character(0),
    font_name = character(0),
    font_size = numeric(0),
    header_bold = logical(0),
    header_color = character(0),
    header_bg = character(0),
    body_bg = character(0),
    alternating_rows = logical(0),
    alternating_bg = character(0),
    padding = numeric(0),
    stringsAsFactors = FALSE
  )
}

load_custom_styles <- function() {
  custom_styles_file <- file.path(PATTERNS_DIR, ".custom_styles.csv")
  if (file.exists(custom_styles_file)) {
    tryCatch({
      df <- read.csv(custom_styles_file, stringsAsFactors = FALSE)
      # Convert logical strings back to logical
      if ("header_bold" %in% names(df)) {
        df$header_bold <- as.logical(df$header_bold)
      }
      if ("alternating_rows" %in% names(df)) {
        df$alternating_rows <- as.logical(df$alternating_rows)
      }
      if ("font_size" %in% names(df)) {
        df$font_size <- as.numeric(df$font_size)
      }
      if ("padding" %in% names(df)) {
        df$padding <- as.numeric(df$padding)
      }
      df
    }, error = function(e) {
      get_empty_custom_styles_df()
    })
  } else {
    get_empty_custom_styles_df()
  }
}

save_custom_styles <- function(df) {
  init_patterns_dir()
  custom_styles_file <- file.path(PATTERNS_DIR, ".custom_styles.csv")
  if (nrow(df) > 0L) {
    write.csv(df, custom_styles_file, row.names = FALSE)
  }
}

delete_custom_style <- function(style_name) {
  df <- load_custom_styles()
  if (style_name %in% df$name) {
    df <- df[df$name != style_name, , drop = FALSE]
    save_custom_styles(df)
    return(TRUE)
  }
  FALSE
}

# ── Custom Styles Management ─────────────────────────────────────────────────
# Store custom table styles that users can define and save

# Default empty custom styles dataframe
get_empty_custom_styles_df <- function() {
  data.frame(
    name = character(0),
    font_name = character(0),
    font_size = numeric(0),
    header_bold = logical(0),
    header_color = character(0),
    header_bg = character(0),
    body_bg = character(0),
    alternating_rows = logical(0),
    alternating_bg = character(0),
    padding = numeric(0),
    stringsAsFactors = FALSE
  )
}

load_custom_styles <- function() {
  custom_styles_file <- file.path(PATTERNS_DIR, ".custom_styles.csv")
  if (file.exists(custom_styles_file)) {
    tryCatch({
      df <- read.csv(custom_styles_file, stringsAsFactors = FALSE)
      # Convert logical strings back to logical
      if ("header_bold" %in% names(df)) {
        df$header_bold <- as.logical(df$header_bold)
      }
      if ("alternating_rows" %in% names(df)) {
        df$alternating_rows <- as.logical(df$alternating_rows)
      }
      if ("font_size" %in% names(df)) {
        df$font_size <- as.numeric(df$font_size)
      }
      if ("padding" %in% names(df)) {
        df$padding <- as.numeric(df$padding)
      }
      df
    }, error = function(e) {
      get_empty_custom_styles_df()
    })
  } else {
    get_empty_custom_styles_df()
  }
}

save_custom_styles <- function(df) {
  init_patterns_dir()
  custom_styles_file <- file.path(PATTERNS_DIR, ".custom_styles.csv")
  if (nrow(df) > 0L) {
    write.csv(df, custom_styles_file, row.names = FALSE)
  }
}

delete_custom_style <- function(style_name) {
  df <- load_custom_styles()
  if (style_name %in% df$name) {
    df <- df[df$name != style_name, , drop = FALSE]
    save_custom_styles(df)
    return(TRUE)
  }
  FALSE
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

# ── Apply formatting rules (regex-based transformations) ────────────────────
# Applies a list of regex-based formatting rules to numeric and numeric-like values
# Only processes cells that are numeric-like to avoid breaking text
apply_formatting_rules <- function(df, rules) {
  if (is.null(rules) || length(rules) == 0L) return(df)

  result <- lapply(seq_len(ncol(df)), function(j) {
    vapply(df[[j]], function(v) {
      if (!nzchar(v) || !is_numeric_like(v)) return(v)
      # Apply each enabled rule in sequence
      result_val <- v
      for (rule in rules) {
        if (isTRUE(rule$enabled)) {
          tryCatch({
            result_val <- gsub(rule$pattern, rule$replacement, result_val, perl = TRUE)
          }, error = function(e) {
            # Silently skip malformed regex patterns
            warning(sprintf("Regex error in rule '%s': %s", rule$name, e$message))
            return(result_val)
          })
        }
      }
      result_val
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
# 3 · TABLE STYLES  (4 styles + custom)
# 3 · TABLE STYLES  (4 styles + custom)
# ══════════════════════════════════════════════════════════════════════════════
# Each style applies non-border properties; borders are applied separately
# so that the title header row (prepended last) does not interfere.

STYLES <- c(
  "Klinisch Standard"  = "clinical",
  "Kompakt"            = "compact",
  "Minimal (APA)"      = "apa",
  "Formales Raster"    = "grid"
)

# Get available styles including custom styles
get_available_styles <- function() {
  custom_df <- load_custom_styles()
  if (nrow(custom_df) > 0L) {
    custom_styles <- setNames(paste0("custom_", seq_len(nrow(custom_df))), custom_df$name)
    c(STYLES, custom_styles)
  } else {
    STYLES
  }
}

# Get style properties from custom styles dataframe
get_custom_style_props <- function(style_name, custom_styles_df) {
  idx <- which(custom_styles_df$name == style_name)
  if (length(idx) == 0L) return(NULL)
  as.list(custom_styles_df[idx[1L], , drop = FALSE])
}

# Get available styles including custom styles
get_available_styles <- function() {
  custom_df <- load_custom_styles()
  if (nrow(custom_df) > 0L) {
    custom_styles <- setNames(paste0("custom_", seq_len(nrow(custom_df))), custom_df$name)
    c(STYLES, custom_styles)
  } else {
    STYLES
  }
}

# Get style properties from custom styles dataframe
get_custom_style_props <- function(style_name, custom_styles_df) {
  idx <- which(custom_styles_df$name == style_name)
  if (length(idx) == 0L) return(NULL)
  as.list(custom_styles_df[idx[1L], , drop = FALSE])
}

# Apply font / colour / alignment / padding – NO borders
.style_props <- function(ft, style_id, custom_styles_df = NULL) {
  nc     <- length(ft$col_keys)
  data_j <- if (nc >= 2L) seq(2L, nc) else integer(0)
  nr_b   <- nrow_part(ft, "body")

  # Check if this is a custom style
  if (!is.null(custom_styles_df) && grepl("^custom_", style_id)) {
    custom_idx <- as.integer(sub("^custom_", "", style_id))
    if (custom_idx > 0L && custom_idx <= nrow(custom_styles_df)) {
      style_row <- custom_styles_df[custom_idx, ]
      
      font_nm <- style_row$font_name
      font_sz <- style_row$font_size
      pad     <- style_row$padding
      pad_lr  <- pad + 1L
      
      ft <- font(ft, fontname = font_nm, part = "all")
      ft <- fontsize(ft, size = font_sz, part = "all")
      ft <- align(ft, align = "left", part = "all")
      ft <- padding(ft, padding.top = pad, padding.bottom = pad,
                    padding.left = pad_lr, padding.right = pad_lr, part = "all")
      ft <- border_remove(ft)
      
      if (length(data_j) > 0) {
        ft <- align(ft, j = data_j, align = "right", part = "body")
        ft <- align(ft, j = data_j, align = "center", part = "header")
      }
      
      # Header styling
      if (isTRUE(style_row$header_bold)) {
        ft <- bold(ft, bold = TRUE, part = "header")
      }
      ft <- color(ft, color = style_row$header_color, part = "header")
      ft <- bg(ft, bg = style_row$header_bg, part = "header")
      
      # Body styling
      ft <- bg(ft, bg = style_row$body_bg, part = "body")
      
      # Alternating rows
      if (isTRUE(style_row$alternating_rows) && nr_b > 0L) {
        odd  <- seq(1L, nr_b, 2L)
        even <- seq(2L, nr_b, 2L)
        if (length(odd))  ft <- bg(ft, i = odd, bg = style_row$alternating_bg, part = "body")
        if (length(even)) ft <- bg(ft, i = even, bg = style_row$body_bg, part = "body")
      }
      
      return(ft)
    }
  }

  # Built-in styles
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

build_flextable <- function(tbl, style_id, lang = "EN", tmap = NULL, fmt_rules = NULL) {

  # Apply translation
  if (lang == "DE" && !is.null(tmap)) {
    tbl$title      <- translate_text(tbl$title, tmap)
    tbl$col_names  <- translate_vec(tbl$col_names, tmap)
    tbl$extra_hdrs <- lapply(tbl$extra_hdrs, translate_vec, tmap = tmap)
    tbl$data       <- translate_df(tbl$data, tmap)
    tbl$footnotes  <- translate_vec(tbl$footnotes, tmap)
    
    # Apply formatting rules (e.g., decimal separator, percentage formatting)
    if (!is.null(fmt_rules)) {
      tbl$data <- apply_formatting_rules(tbl$data, fmt_rules)
    }
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
  custom_styles_df <- load_custom_styles()
  ft <- .style_props(ft, style_id, custom_styles_df)
  custom_styles_df <- load_custom_styles()
  ft <- .style_props(ft, style_id, custom_styles_df)

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
  tags$head(
    tags$script(HTML("
      Shiny.addCustomMessageHandler('refresh_file_list', function(message) {
        // Trigger DataTable refresh by clicking on the tab
        var mainTabs = document.querySelector('[data-value=\"Tabellen & Vorschau\"]');
        if (mainTabs) {
          Shiny.setInputValue('main_tabs', 'Tabellen & Vorschau', {priority: 'event'});
        }
      });
    ")),
    tags$style(HTML("
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
    .modal-xl { width: 100vw !important; }
    .modal-xl .modal-dialog { max-width: 98vw; width: 98vw; max-height: 95vh; }
    .modal-xl .modal-content { max-height: 90vh; overflow-y: auto; }
    .cmp-panel  { overflow-y:auto; max-height:85vh; }
    .modal-xl { width: 100vw !important; }
    .modal-xl .modal-dialog { max-width: 98vw; width: 98vw; max-height: 95vh; }
    .modal-xl .modal-content { max-height: 90vh; overflow-y: auto; }
    .cmp-panel  { overflow-y:auto; max-height:85vh; }
    .cmp-no-pdf { color:#c0392b; font-size:12px; padding:20px;
                  text-align:center; background:#fff5f5; border-radius:4px; }
    /* Progress bar styling */
    .shiny-progress-container { position: fixed; bottom: 20px; right: 20px; 
                                width: 350px; background: white; border-radius: 8px;
                                box-shadow: 0 4px 12px rgba(0,0,0,0.15);
                                padding: 0; z-index: 9999; }
    .shiny-progress { border-radius: 8px; }
    .progress { background-color: #e8e8e8; border-radius: 6px; height: 24px;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin: 12px; }
    .progress-bar { background: linear-gradient(90deg, #27ae60 0%, #2ecc71 100%);
                    height: 100%; border-radius: 6px; font-size: 12px; 
                    color: white; font-weight: 600; display: flex;
                    align-items: center; justify-content: center;
                    transition: width 0.3s ease; }
    .shiny-progress-detail { padding: 0 12px 8px 12px; font-size: 12px;
                             color: #555; text-align: center; font-weight: 500; }
    .shiny-progress-message { padding: 12px 12px 4px 12px; font-size: 13px;
                              font-weight: 600; color: #2c3e50; }
  ")),

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
      uiOutput("style_selector_ui"),
      uiOutput("style_selector_ui"),

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

        # ── Tab 0: Help ────────────────────────────────────────────────────
        tabPanel(
          "Hilfe",
          br(),
          div(style = "max-width:900px; line-height:1.8; color:#333;",
            h3("RTF Tabellen-Generator – Benutzerhandbuch", style = "color:#2c3e50;"),
            
            h4("📋 Übersicht"),
            p("Diese Anwendung vereinfacht die Konvertierung von PDF-Tabellen zu RTF-Format mit ",
              "optionalen Übersetzungs-, Formatierungs- und Stilierungsoptionen. Sie können mehrere PDFs ",
              "verarbeiten, Einstellungen als wiederverwendbare Muster speichern und Batch-Exporte durchführen."),
            
            h4("🚀 Schnelleinstieg"),
            tags$ol(
              tags$li(
                strong("PDFs laden:"), " Gehen Sie zur Registerkarte 'PDFs laden' und laden Sie eine oder mehrere PDF-Dateien hoch. ",
                "Die App extrahiert automatisch alle Tabellen und speichert sie als Excel-Dateien in den Ordner ",
                code("extracted_tables/"), "."
              ),
              tags$li(
                strong("Tabelle auswählen:"), " Wählen Sie eine Tabelle aus der Liste in der Registerkarte 'Tabellen & Vorschau' aus."
              ),
              tags$li(
                strong("Sprache und Stil:"), " Wählen Sie einen Tabellenstil (links in der Seitenleiste) und die Sprache (EN oder DE)."
              ),
              tags$li(
                strong("Download:"), " Klicken Sie auf 'Aktive Tabelle (RTF)' zum Herunterladen einer einzelnen Tabelle ",
                "oder verwenden Sie die Registerkarte 'Batch-Export' für mehrere Tabellen."
              )
            ),
            
            hr(),
            h4("📂 Registerkarten Erklärung"),
            
            tags$div(style = "background:#f8f9fa; padding:12px; border-radius:4px; margin:10px 0;",
              tags$strong("PDFs laden"),
              p("Laden Sie eine oder mehrere PDF-Dateien hoch. Der System führt das Python-Skript ",
                code("extract_tables.py"), " aus, um alle Tabellen automatisch zu extrahieren.")
            ),
            
            tags$div(style = "background:#f8f9fa; padding:12px; border-radius:4px; margin:10px 0;",
              tags$strong("Tabellen & Vorschau"),
              p("Durchsuchen und wählen Sie Tabellen aus der extrahierten Liste aus. ",
                "Eine Live-Vorschau zeigt die formatierte Tabelle mit den aktuellen Einstellungen.")
            ),
            
            tags$div(style = "background:#f8f9fa; padding:12px; border-radius:4px; margin:10px 0;",
              tags$strong("Übersetzungen"),
              p("Zwei Unterkarten: 'Character' für Wort- und Phrasübersetzungen (Englisch → Deutsch) ",
                "und 'Numeric' für reguläre Ausdrücke bei Zahlenformatierung (z. B. Dezimaltrennzeichen).")
            ),
            
            tags$div(style = "background:#f8f9fa; padding:12px; border-radius:4px; margin:10px 0;",
              tags$strong("Muster (Vorlagen)"),
              p("Speichern Sie Ihre aktuellen Einstellungen (Übersetzungen, Formatierungsregeln, Stil) als ",
                "wiederverwendbare Vorlagen für zukünftige Projekte.")
            ),
            
            tags$div(style = "background:#f8f9fa; padding:12px; border-radius:4px; margin:10px 0;",
              tags$strong("Tabellenstile"),
              p("Erstellen Sie benutzerdefinierte Tabellenstile mit Schrift-, Farb- und Abstands-Optionen. ",
                "Speichern Sie Stile global oder speichern Sie sie als Teil eines Musters.")
            ),
            
            tags$div(style = "background:#f8f9fa; padding:12px; border-radius:4px; margin:10px 0;",
              tags$strong("Batch-Export"),
              p("Exportieren Sie mehrere Tabellen auf einmal als ZIP-Datei mit konsistenter Formatierung.")
            ),
            
            hr(),
            h4("⚙️ Abhängigkeiten"),
            
            tags$div(style = "background:#fff3cd; padding:12px; border-radius:4px; margin:10px 0;",
              p(
                tags$strong("Python-Pakete:"), br(),
                "Diese App benötigt Python 3 mit den Paketen ", 
                code("pdfplumber"), " und ", code("openpyxl"), ". ",
                "Sie werden automatisch beim Start überprüft und installiert. "
              )
            ),
            
            tags$div(style = "background:#fff3cd; padding:12px; border-radius:4px; margin:10px 0;",
              p(
                tags$strong("Optional für PDF-Vergleich:"), br(),
                "Um die Registerkarte 'PDF vs RTF' verwenden zu können, benötigen Sie:", br(),
                code("brew install poppler"), br(),
                "und dann in R:", br(),
                code("install.packages(c('pdftools', 'png'))")
              )
            ),
            
            hr(),
            h4("💡 Tipps & Tricks"),
            tags$ul(
              tags$li("Nutzen Sie die Suchfunktion in der Tabellenliste, um schnell Tabellen zu finden."),
              tags$li("Speichern Sie häufig verwendete Einstellungen als Muster für wiederverwendbaren Zugriff."),
              tags$li("Sie können Übersetzungen und Formatierungsregeln direkt in den Tabellen bearbeiten."),
              tags$li("Der Batch-Export erzeugt eine ZIP-Datei mit allen ausgewählten Tabellen."),
              tags$li("Muster werden im Ordner 'muster/' gespeichert und können projektübergreifend verwendet werden.")
            ),
            
            hr(),
            h4("❓ Häufig gestellte Fragen"),
            
            tags$div(style = "background:#e7f3ff; padding:12px; border-radius:4px; margin:10px 0;",
              tags$strong("F: Wie lade ich eine neue PDF-Datei?"),
              p("A: Gehen Sie zur Registerkarte 'PDFs laden', klicken Sie auf 'Datei auswählen' und wählen Sie ",
                "eine oder mehrere PDF-Dateien. Die App extrahiert alle Tabellen automatisch.")
            ),
            
            tags$div(style = "background:#e7f3ff; padding:12px; border-radius:4px; margin:10px 0;",
              tags$strong("F: Kann ich mehrere PDFs gleichzeitig verarbeiten?"),
              p("A: Ja! Wählen Sie einfach mehrere PDFs in der Dateiauswahl und laden Sie sie gleichzeitig hoch.")
            ),
            
            tags$div(style = "background:#e7f3ff; padding:12px; border-radius:4px; margin:10px 0;",
              tags$strong("F: Wie speichere ich meine Einstellungen?"),
              p("A: Verwenden Sie die Registerkarte 'Muster' zur Speicherung. Klicken Sie auf 'Muster speichern' und geben Sie einen Namen ein.")
            ),
            
            tags$div(style = "background:#e7f3ff; padding:12px; border-radius:4px; margin:10px 0;",
              tags$strong("F: Was ist der Unterschied zwischen 'Aktive Tabelle' und 'Batch-Export'?"),
              p("A: 'Aktive Tabelle' exportiert nur eine Tabelle als RTF. 'Batch-Export' erlaubt die Auswahl mehrerer ",
                "Tabellen und erzeugt eine ZIP-Datei mit allen ausgewählten Tabellen.")
            )
          )
        ),

        # ── Tab 1: PDFs laden ──────────────────────────────────────────────
        tabPanel(
          "PDFs laden",
          br(),
          fluidRow(
            column(8,
              wellPanel(
                div(class = "section-lbl", "PDF-Dateien hochladen"),
                fileInput("pdf_upload", "Wählen Sie eine oder mehrere PDF-Dateien:",
                          accept = c(".pdf"),
                          multiple = TRUE,
                          width = "100%"),
                helpText(style = "font-size:11px; color:#666; margin-top:6px;",
                         "Sie können mehrere PDFs gleichzeitig hochladen. Die Verarbeitung kann bei großen PDFs einige Zeit dauern."),
                div(style = "margin-top:12px;",
                  actionButton("btn_extract", "Tabellen extrahieren",
                               icon = icon("play"),
                               class = "btn-primary",
                               style = "width:100%;")
                )
              )
            ),
            column(4,
              wellPanel(
                div(class = "section-lbl", "Verarbeitungsstatus"),
                div(class = "info-box",
                    verbatimTextOutput("extract_status"))
              )
            )
          ),
          br(),
          fluidRow(
            column(12,
              wellPanel(
                div(class = "section-lbl", "Extrahierte Tabellen"),
                div(class = "info-box",
                    verbatimTextOutput("extraction_summary"))
              )
            )
          ),
          br(),
          fluidRow(
            column(12,
              wellPanel(
                div(class = "section-lbl", "Abhängigkeitsprüfung"),
                div(class = "info-box",
                    htmlOutput("dependency_status"))
              )
            )
          )
        ),

        # ── Tab 2: Browser + Vorschau ──────────────────────────────────────
        # ── Tab 0: Help ────────────────────────────────────────────────────
        tabPanel(
          "Hilfe",
          br(),
          div(style = "max-width:900px; line-height:1.8; color:#333;",
            h3("RTF Tabellen-Generator – Benutzerhandbuch", style = "color:#2c3e50;"),
            
            h4("📋 Übersicht"),
            p("Diese Anwendung vereinfacht die Konvertierung von PDF-Tabellen zu RTF-Format mit ",
              "optionalen Übersetzungs-, Formatierungs- und Stilierungsoptionen. Sie können mehrere PDFs ",
              "verarbeiten, Einstellungen als wiederverwendbare Muster speichern und Batch-Exporte durchführen."),
            
            h4("🚀 Schnelleinstieg"),
            tags$ol(
              tags$li(
                strong("PDFs laden:"), " Gehen Sie zur Registerkarte 'PDFs laden' und laden Sie eine oder mehrere PDF-Dateien hoch. ",
                "Die App extrahiert automatisch alle Tabellen und speichert sie als Excel-Dateien in den Ordner ",
                code("extracted_tables/"), "."
              ),
              tags$li(
                strong("Tabelle auswählen:"), " Wählen Sie eine Tabelle aus der Liste in der Registerkarte 'Tabellen & Vorschau' aus."
              ),
              tags$li(
                strong("Sprache und Stil:"), " Wählen Sie einen Tabellenstil (links in der Seitenleiste) und die Sprache (EN oder DE)."
              ),
              tags$li(
                strong("Download:"), " Klicken Sie auf 'Aktive Tabelle (RTF)' zum Herunterladen einer einzelnen Tabelle ",
                "oder verwenden Sie die Registerkarte 'Batch-Export' für mehrere Tabellen."
              )
            ),
            
            hr(),
            h4("📂 Registerkarten Erklärung"),
            
            tags$div(style = "background:#f8f9fa; padding:12px; border-radius:4px; margin:10px 0;",
              tags$strong("PDFs laden"),
              p("Laden Sie eine oder mehrere PDF-Dateien hoch. Der System führt das Python-Skript ",
                code("extract_tables.py"), " aus, um alle Tabellen automatisch zu extrahieren.")
            ),
            
            tags$div(style = "background:#f8f9fa; padding:12px; border-radius:4px; margin:10px 0;",
              tags$strong("Tabellen & Vorschau"),
              p("Durchsuchen und wählen Sie Tabellen aus der extrahierten Liste aus. ",
                "Eine Live-Vorschau zeigt die formatierte Tabelle mit den aktuellen Einstellungen.")
            ),
            
            tags$div(style = "background:#f8f9fa; padding:12px; border-radius:4px; margin:10px 0;",
              tags$strong("Übersetzungen"),
              p("Zwei Unterkarten: 'Character' für Wort- und Phrasübersetzungen (Englisch → Deutsch) ",
                "und 'Numeric' für reguläre Ausdrücke bei Zahlenformatierung (z. B. Dezimaltrennzeichen).")
            ),
            
            tags$div(style = "background:#f8f9fa; padding:12px; border-radius:4px; margin:10px 0;",
              tags$strong("Muster (Vorlagen)"),
              p("Speichern Sie Ihre aktuellen Einstellungen (Übersetzungen, Formatierungsregeln, Stil) als ",
                "wiederverwendbare Vorlagen für zukünftige Projekte.")
            ),
            
            tags$div(style = "background:#f8f9fa; padding:12px; border-radius:4px; margin:10px 0;",
              tags$strong("Tabellenstile"),
              p("Erstellen Sie benutzerdefinierte Tabellenstile mit Schrift-, Farb- und Abstands-Optionen. ",
                "Speichern Sie Stile global oder speichern Sie sie als Teil eines Musters.")
            ),
            
            tags$div(style = "background:#f8f9fa; padding:12px; border-radius:4px; margin:10px 0;",
              tags$strong("Batch-Export"),
              p("Exportieren Sie mehrere Tabellen auf einmal als ZIP-Datei mit konsistenter Formatierung.")
            ),
            
            hr(),
            h4("⚙️ Abhängigkeiten"),
            
            tags$div(style = "background:#fff3cd; padding:12px; border-radius:4px; margin:10px 0;",
              p(
                tags$strong("Python-Pakete:"), br(),
                "Diese App benötigt Python 3 mit den Paketen ", 
                code("pdfplumber"), " und ", code("openpyxl"), ". ",
                "Sie werden automatisch beim Start überprüft und installiert. "
              )
            ),
            
            tags$div(style = "background:#fff3cd; padding:12px; border-radius:4px; margin:10px 0;",
              p(
                tags$strong("Optional für PDF-Vergleich:"), br(),
                "Um die Registerkarte 'PDF vs RTF' verwenden zu können, benötigen Sie:", br(),
                code("brew install poppler"), br(),
                "und dann in R:", br(),
                code("install.packages(c('pdftools', 'png'))")
              )
            ),
            
            hr(),
            h4("💡 Tipps & Tricks"),
            tags$ul(
              tags$li("Nutzen Sie die Suchfunktion in der Tabellenliste, um schnell Tabellen zu finden."),
              tags$li("Speichern Sie häufig verwendete Einstellungen als Muster für wiederverwendbaren Zugriff."),
              tags$li("Sie können Übersetzungen und Formatierungsregeln direkt in den Tabellen bearbeiten."),
              tags$li("Der Batch-Export erzeugt eine ZIP-Datei mit allen ausgewählten Tabellen."),
              tags$li("Muster werden im Ordner 'muster/' gespeichert und können projektübergreifend verwendet werden.")
            ),
            
            hr(),
            h4("❓ Häufig gestellte Fragen"),
            
            tags$div(style = "background:#e7f3ff; padding:12px; border-radius:4px; margin:10px 0;",
              tags$strong("F: Wie lade ich eine neue PDF-Datei?"),
              p("A: Gehen Sie zur Registerkarte 'PDFs laden', klicken Sie auf 'Datei auswählen' und wählen Sie ",
                "eine oder mehrere PDF-Dateien. Die App extrahiert alle Tabellen automatisch.")
            ),
            
            tags$div(style = "background:#e7f3ff; padding:12px; border-radius:4px; margin:10px 0;",
              tags$strong("F: Kann ich mehrere PDFs gleichzeitig verarbeiten?"),
              p("A: Ja! Wählen Sie einfach mehrere PDFs in der Dateiauswahl und laden Sie sie gleichzeitig hoch.")
            ),
            
            tags$div(style = "background:#e7f3ff; padding:12px; border-radius:4px; margin:10px 0;",
              tags$strong("F: Wie speichere ich meine Einstellungen?"),
              p("A: Verwenden Sie die Registerkarte 'Muster' zur Speicherung. Klicken Sie auf 'Muster speichern' und geben Sie einen Namen ein.")
            ),
            
            tags$div(style = "background:#e7f3ff; padding:12px; border-radius:4px; margin:10px 0;",
              tags$strong("F: Was ist der Unterschied zwischen 'Aktive Tabelle' und 'Batch-Export'?"),
              p("A: 'Aktive Tabelle' exportiert nur eine Tabelle als RTF. 'Batch-Export' erlaubt die Auswahl mehrerer ",
                "Tabellen und erzeugt eine ZIP-Datei mit allen ausgewählten Tabellen.")
            )
          )
        ),

        # ── Tab 1: PDFs laden ──────────────────────────────────────────────
        tabPanel(
          "PDFs laden",
          br(),
          fluidRow(
            column(8,
              wellPanel(
                div(class = "section-lbl", "PDF-Dateien hochladen"),
                fileInput("pdf_upload", "Wählen Sie eine oder mehrere PDF-Dateien:",
                          accept = c(".pdf"),
                          multiple = TRUE,
                          width = "100%"),
                helpText(style = "font-size:11px; color:#666; margin-top:6px;",
                         "Sie können mehrere PDFs gleichzeitig hochladen. Die Verarbeitung kann bei großen PDFs einige Zeit dauern."),
                div(style = "margin-top:12px;",
                  actionButton("btn_extract", "Tabellen extrahieren",
                               icon = icon("play"),
                               class = "btn-primary",
                               style = "width:100%;")
                )
              )
            ),
            column(4,
              wellPanel(
                div(class = "section-lbl", "Verarbeitungsstatus"),
                div(class = "info-box",
                    verbatimTextOutput("extract_status"))
              )
            )
          ),
          br(),
          fluidRow(
            column(12,
              wellPanel(
                div(class = "section-lbl", "Extrahierte Tabellen"),
                div(class = "info-box",
                    verbatimTextOutput("extraction_summary"))
              )
            )
          ),
          br(),
          fluidRow(
            column(12,
              wellPanel(
                div(class = "section-lbl", "Abhängigkeitsprüfung"),
                div(class = "info-box",
                    htmlOutput("dependency_status"))
              )
            )
          )
        ),

        # ── Tab 2: Browser + Vorschau ──────────────────────────────────────
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

        # ── Tab 3: Übersetzungen ───────────────────────────────────────────
        # ── Tab 3: Übersetzungen ───────────────────────────────────────────
        tabPanel(
          "Übersetzungen",
          br(),
          tabsetPanel(
            # ── Subtab 2.1: Character translations ──────────────────────────
            tabPanel(
              "Character",
              br(),
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
            ),
            
            # ── Subtab 2.2: Numeric formatting rules ────────────────────────
            tabPanel(
              "Numeric",
              br(),
              div(style = "margin-bottom:10px;",
                actionButton("btn_save_fmt_rules",  "Speichern",
                             icon = icon("floppy-disk"),
                             class = "btn-success btn-sm"),
                actionButton("btn_reset_fmt_rules", "Zurücksetzen",
                             icon = icon("rotate-left"),
                             class = "btn-warning btn-sm",
                             style = "margin-left:6px;"),
                actionButton("btn_add_fmt_rule",    "Regel hinzufügen",
                             icon = icon("plus"),
                             class = "btn-info btn-sm",
                             style = "margin-left:6px;"),
                actionButton("btn_del_fmt_rules",   "Auswahl löschen",
                             icon = icon("trash"),
                             class = "btn-danger btn-sm",
                             style = "margin-left:6px;")
              ),
              helpText(style = "font-size:11px; color:#666;",
                       "Doppelklick auf eine Zelle zum Bearbeiten. ",
                       "Pattern: Regex für die Suche (z.B. '\\\\.') | ",
                       "Replacement: Ersetzungstext (z.B. ',') | ",
                       "Enabled: Aktiviert/Deaktiviert die Regel | ",
                       "\"Speichern\" schreibt in formatting_rules_de.csv."),
              DTOutput("fmt_rules_tbl")
            )
          )
        ),

        # ── Tab 4: Muster (Pattern Templates) ──────────────────────────────
        # ── Tab 4: Muster (Pattern Templates) ──────────────────────────────
        tabPanel(
          "Muster",
          br(),
          fluidRow(
            column(6,
              wellPanel(
                div(class = "section-lbl", "Muster speichern"),
                textInput("pattern_save_name", "Muster Name:",
                          placeholder = "z.B. 'Klinische Studie 2024'"),
                actionButton("btn_save_pattern", "Muster speichern",
                             icon = icon("floppy-disk"),
                             class = "btn-success"),
                helpText(style = "font-size:11px; color:#666; margin-top:6px;",
                         "Speichert Übersetzungen, Formatierungsregeln und ",
                         "Tabellenstil als Muster für später Verwendung."),
                hr(),
                div(class = "section-lbl", "Verfügbare Muster laden"),
                uiOutput("pattern_list_ui"),
                div(style = "margin-top:8px;",
                  actionButton("btn_load_pattern", "Muster laden",
                               icon = icon("download"),
                               class = "btn-primary btn-sm"),
                  actionButton("btn_delete_pattern", "Muster löschen",
                               icon = icon("trash"),
                               class = "btn-danger btn-sm",
                               style = "margin-left:4px;")
                ),
                helpText(style = "font-size:11px; color:#666; margin-top:6px;",
                         "Muster werden im Ordner 'muster' gespeichert.")
              )
            ),
            column(6,
              wellPanel(
                div(class = "section-lbl", "Muster Informationen"),
                div(class = "info-box",
                    verbatimTextOutput("pattern_info_text")
                )
              )
            )
          )
        ),

        # ── Tab 5: Tabellenstile (Custom Styles) ──────────────────────────
        tabPanel(
          "Tabellenstile",
          br(),
          fluidRow(
            column(12,
              wellPanel(
                div(class = "section-lbl", "Neuer Tabellenstil"),
                textInput("custom_style_name", "Stilname:",
                          placeholder = "z.B. 'Mein blauer Stil'"),
                div(style = "margin-top:12px;",
                  column(3,
                    div(class = "section-lbl", "Schrift"),
                    selectInput("cs_font_name", "Schriftart:",
                                choices = c("Arial", "Times New Roman", "Courier", "Verdana"),
                                selected = "Arial", width = "100%"),
                    numericInput("cs_font_size", "Größe:", value = 9, min = 6, max = 16, step = 1, width = "100%")
                  ),
                  column(3,
                    div(class = "section-lbl", "Kopfzeile"),
                    checkboxInput("cs_header_bold", "Fett", value = TRUE),
                    colourpicker::colourInput("cs_header_color", "Textfarbe:", value = "#ffffff", 
                                              palette = "limited", width = "100%"),
                    colourpicker::colourInput("cs_header_bg", "Hintergrund:", value = "#1a3a5c",
                                              palette = "limited", width = "100%")
                  ),
                  column(3,
                    div(class = "section-lbl", "Datenzellen"),
                    colourpicker::colourInput("cs_body_bg", "Hintergrund:", value = "#ffffff",
                                              palette = "limited", width = "100%"),
                    checkboxInput("cs_alternating", "Wechselnde Zeilen", value = FALSE),
                    conditionalPanel(
                      condition = "input.cs_alternating",
                      colourpicker::colourInput("cs_alternating_bg", "Alt. Hintergrund:", 
                                                value = "#f0f5fb", palette = "limited", width = "100%")
                    )
                  ),
                  column(3,
                    div(class = "section-lbl", "Abstände"),
                    numericInput("cs_padding", "Padding (pt):", value = 3, min = 1, max = 10, step = 1, width = "100%")
                  )
                ),
                div(style = "margin-top:12px;",
                  actionButton("btn_preview_custom_style", "Vorschau",
                               icon = icon("eye"),
                               class = "btn-info btn-sm"),
                  actionButton("btn_save_custom_style", "Stil speichern",
                               icon = icon("floppy-disk"),
                               class = "btn-success btn-sm",
                               style = "margin-left:6px;"),
                  actionButton("btn_delete_custom_style", "Stil löschen",
                               icon = icon("trash"),
                               class = "btn-danger btn-sm",
                               style = "margin-left:6px;")
                )
              )
            )
          ),
          br(),
          fluidRow(
            column(12,
              wellPanel(
                div(class = "section-lbl", "Gespeicherte Tabellenstile"),
                DTOutput("custom_styles_tbl")
              )
            )
          ),
          br(),
          fluidRow(
            column(12,
              wellPanel(
                div(class = "section-lbl", "Vorschau"),
                div(class = "preview-box",
                    uiOutput("custom_style_preview"))
              )
            )
          )
        ),

        # ── Tab 6: Batch-Export ────────────────────────────────────────────
        # ── Tab 5: Tabellenstile (Custom Styles) ──────────────────────────
        tabPanel(
          "Tabellenstile",
          br(),
          fluidRow(
            column(12,
              wellPanel(
                div(class = "section-lbl", "Neuer Tabellenstil"),
                textInput("custom_style_name", "Stilname:",
                          placeholder = "z.B. 'Mein blauer Stil'"),
                div(style = "margin-top:12px;",
                  column(3,
                    div(class = "section-lbl", "Schrift"),
                    selectInput("cs_font_name", "Schriftart:",
                                choices = c("Arial", "Times New Roman", "Courier", "Verdana"),
                                selected = "Arial", width = "100%"),
                    numericInput("cs_font_size", "Größe:", value = 9, min = 6, max = 16, step = 1, width = "100%")
                  ),
                  column(3,
                    div(class = "section-lbl", "Kopfzeile"),
                    checkboxInput("cs_header_bold", "Fett", value = TRUE),
                    colourpicker::colourInput("cs_header_color", "Textfarbe:", value = "#ffffff", 
                                              palette = "limited", width = "100%"),
                    colourpicker::colourInput("cs_header_bg", "Hintergrund:", value = "#1a3a5c",
                                              palette = "limited", width = "100%")
                  ),
                  column(3,
                    div(class = "section-lbl", "Datenzellen"),
                    colourpicker::colourInput("cs_body_bg", "Hintergrund:", value = "#ffffff",
                                              palette = "limited", width = "100%"),
                    checkboxInput("cs_alternating", "Wechselnde Zeilen", value = FALSE),
                    conditionalPanel(
                      condition = "input.cs_alternating",
                      colourpicker::colourInput("cs_alternating_bg", "Alt. Hintergrund:", 
                                                value = "#f0f5fb", palette = "limited", width = "100%")
                    )
                  ),
                  column(3,
                    div(class = "section-lbl", "Abstände"),
                    numericInput("cs_padding", "Padding (pt):", value = 3, min = 1, max = 10, step = 1, width = "100%")
                  )
                ),
                div(style = "margin-top:12px;",
                  actionButton("btn_preview_custom_style", "Vorschau",
                               icon = icon("eye"),
                               class = "btn-info btn-sm"),
                  actionButton("btn_save_custom_style", "Stil speichern",
                               icon = icon("floppy-disk"),
                               class = "btn-success btn-sm",
                               style = "margin-left:6px;"),
                  actionButton("btn_delete_custom_style", "Stil löschen",
                               icon = icon("trash"),
                               class = "btn-danger btn-sm",
                               style = "margin-left:6px;")
                )
              )
            )
          ),
          br(),
          fluidRow(
            column(12,
              wellPanel(
                div(class = "section-lbl", "Gespeicherte Tabellenstile"),
                DTOutput("custom_styles_tbl")
              )
            )
          ),
          br(),
          fluidRow(
            column(12,
              wellPanel(
                div(class = "section-lbl", "Vorschau"),
                div(class = "preview-box",
                    uiOutput("custom_style_preview"))
              )
            )
          )
        ),

        # ── Tab 6: Batch-Export ────────────────────────────────────────────
        tabPanel(
          "Batch-Export",
          br(),
          fluidRow(
            column(4,
              wellPanel(
                div(class = "section-lbl", "Batch-Optionen"),
                uiOutput("batch_style_selector_ui"),
                uiOutput("batch_style_selector_ui"),
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
        ),

        # ── Tab 7: Combine RTFs with TOC ──────────────────────────────────
        tabPanel(
          "RTFs kombinieren",
          br(),
          fluidRow(
            column(4,
              wellPanel(
                div(class = "section-lbl", "Einstellungen"),
                textInput("combined_rtf_title", "Dokumenttitel:",
                          placeholder = "z.B. 'Klinische Studienergebnisse'",
                          value = "Klinische Studienergebnisse"),
                hr(),
                div(class = "section-lbl", "Tabellen auswählen"),
                div(style = "margin-bottom:6px;",
                  actionButton("comb_all",  "Alle wählen",
                               class = "btn-sm btn-default"),
                  actionButton("comb_none", "Alle abwählen",
                               class = "btn-sm btn-default",
                               style = "margin-left:4px;")
                ),
                uiOutput("comb_sel_info"),
                hr(),
                checkboxInput("comb_include_toc", "Inhaltsverzeichnis einschließen", 
                              value = TRUE),
                checkboxInput("comb_include_bookmarks", "Lesezeichen hinzufügen", 
                              value = TRUE),
                checkboxInput("comb_page_breaks", "Seitenumbrüche zwischen Tabellen", 
                              value = TRUE),
                hr(),
                downloadButton("dl_combined_rtf", "RTF kombinieren & herunterladen",
                               class = "btn-success",
                               style = "width:100%;"),
                helpText(style = "font-size:11px; color:#888; margin-top:6px;",
                         "Kombiniert ausgewählte Tabellen in eine RTF-Datei mit Inhaltsverzeichnis und Lesezeichen.")
              )
            ),
            column(8,
              div(class = "section-lbl", "Tabellen auswählen:"),
              DTOutput("comb_tbl")
            )
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

  # Initialize patterns directory
  init_patterns_dir()

  # ── Dependency Status ──────────────────────────────────────────────────────
  output$dependency_status <- renderUI({
    py_status <- python_env_status
    
    # Check R packages
    r_check <- all(sapply(REQUIRED_R_PACKAGES, requireNamespace, quietly = TRUE))
    
    html_content <- ""
    
    # R Packages Status
    html_content <- paste0(html_content,
      "<strong style='color:#2c3e50;'>R-Pakete:</strong><br/>",
      if (r_check) {
        "<span style='color:#27ae60;'>✓ Alle erforderlich Pakete installiert</span>"
      } else {
        "<span style='color:#e74c3c;'>✗ Einige R-Pakete fehlen</span>"
      }, "<br/><br/>"
    )
    
    # Python Status
    if (py_status$python_available) {
      html_content <- paste0(html_content,
        "<strong style='color:#2c3e50;'>Python:</strong><br/>",
        "<span style='color:#27ae60;'>✓ Python verfügbar: ", py_status$python_version, "</span><br/>"
      )
      
      if (py_status$all_packages_available) {
        html_content <- paste0(html_content,
          "<span style='color:#27ae60;'>✓ Alle Python-Pakete installiert:</span><br/>",
          "   • pdfplumber<br/>",
          "   • openpyxl<br/>"
        )
      } else {
        html_content <- paste0(html_content,
          "<span style='color:#f39c12;'>⚠ Einige Python-Pakete fehlen:</span><br/>",
          paste0("   • ", py_status$missing_packages, "<br/>", collapse = "")
        )
      }
    } else {
      html_content <- paste0(html_content,
        "<strong style='color:#2c3e50;'>Python:</strong><br/>",
        "<span style='color:#e74c3c;'>✗ Python 3 nicht gefunden. Bitte installieren Sie Python 3.</span><br/>"
      )
    }
    
    # Optional pdftools
    html_content <- paste0(html_content,
      "<br/><strong style='color:#2c3e50;'>Optional (PDF-Vergleich):</strong><br/>",
      if (HAS_PDFTOOLS) {
        "<span style='color:#27ae60;'>✓ pdftools und png installiert</span>"
      } else {
        "<span style='color:#f39c12;'>⚠ pdftools/png nicht installiert (PDF-Vergleich deaktiviert)</span>"
      }
    )
    
    HTML(html_content)
  })

  # ── PDF Upload and Extraction ──────────────────────────────────────────────
  extract_status_msg <- reactiveVal("Bereit für PDF-Upload")
  extraction_summary_msg <- reactiveVal("Keine Extraktion durchgeführt")
  
  output$extract_status <- renderText({
    extract_status_msg()
  })
  
  output$extraction_summary <- renderText({
    extraction_summary_msg()
  })
  
  observeEvent(input$btn_extract, {
    req(input$pdf_upload)
    
    # Check Python availability
    if (!python_env_status$python_available) {
      extract_status_msg("Fehler: Python 3 nicht verfügbar")
      return()
    }
    
    # Install missing Python packages if needed
    if (!python_env_status$all_packages_available) {
      extract_status_msg("Installiere fehlende Python-Pakete...")
      tryCatch({
        install_python_packages(python_env_status$missing_packages)
        extract_status_msg("Python-Pakete installiert. Starte Extraktion...")
      }, error = function(e) {
        extract_status_msg(paste("Fehler beim Installieren von Python-Paketen:", e$message))
        return()
      })
    }
    
    # Create temp directory for uploads
    upload_dir <- tempfile(pattern = "pdf_upload_")
    dir.create(upload_dir, recursive = TRUE)
    
    # Copy uploaded files to temp directory
    pdf_files <- character(0)
    for (i in seq_len(nrow(input$pdf_upload))) {
      src <- input$pdf_upload$datapath[i]
      dst <- file.path(upload_dir, input$pdf_upload$name[i])
      file.copy(src, dst)
      pdf_files <- c(pdf_files, dst)
    }
    
    # Extract tables from each PDF with progress bar
    total_tables <- 0
    errors <- character(0)
    n_pdfs <- length(pdf_files)
    
    # Use withProgress for visual feedback
    withProgress(
      message = "📊 Tabellen extrahieren",
      detail = sprintf("0 / %d PDFs verarbeitet", n_pdfs),
      value = 0,
      {
        for (idx in seq_along(pdf_files)) {
          pdf_file <- pdf_files[idx]
          pdf_name <- basename(pdf_file)
          
          # Update progress message
          incProgress(
            amount = 0,
            detail = sprintf(
              "(%d / %d) Datei: %s",
              idx,
              n_pdfs,
              pdf_name
            )
          )
          
          # Run Python extraction script
          tryCatch({
            cmd_output <- system2(
              "python3",
              c("extract_tables.py", pdf_file, "extracted_tables"),
              stdout = TRUE,
              stderr = TRUE,
              cwd = getwd()
            )
            
            # Count extracted tables from output
            table_lines <- grep("→", cmd_output, value = TRUE)
            num_tables <- length(table_lines)
            total_tables <- total_tables + num_tables
            
            # Update status message with extraction details
            extract_status_msg(
              sprintf(
                "✓ %s\n  Tabellen: %d | Gesamt: %d",
                pdf_name,
                num_tables,
                total_tables
              )
            )
            
          }, error = function(e) {
            errors <<- c(errors, sprintf("%s: %s", pdf_name, e$message))
            extract_status_msg(
              sprintf(
                "✗ %s\n  Fehler: %s",
                pdf_name,
                e$message
              )
            )
          })
          
          # Increment progress bar
          incProgress(
            amount = 1 / n_pdfs,
            detail = sprintf(
              "(%d / %d) %d Tabelle(n) gefunden",
              idx,
              n_pdfs,
              total_tables
            )
          )
        }
      }
    )
    
    # Update UI with final results
    if (length(errors) > 0L) {
      summary_text <- sprintf(
        "⚠️ Extraktion abgeschlossen mit Fehlern.\n\nInsgesamt %d Tabelle(n) extrahiert.\n\n🔴 Fehler:\n%s",
        total_tables,
        paste(errors, collapse = "\n")
      )
      extract_status_msg(sprintf("⚠️ %d Tabelle(n) extrahiert mit %d Fehler(n)", 
                                 total_tables, length(errors)))
    } else {
      summary_text <- sprintf(
        "✅ Extraktion erfolgreich abgeschlossen!\n\n🎉 Insgesamt %d Tabelle(n) extrahiert.\n\n📁 Ort: extracted_tables/\n\n→ Jetzt zur 'Tabellen & Vorschau' Tab wechseln um die Tabellen zu verarbeiten.",
        total_tables
      )
      extract_status_msg(sprintf("✅ %d Tabelle(n) erfolgreich extrahiert", total_tables))
    }
    
    extraction_summary_msg(summary_text)
    
    # Refresh file list
    session$sendCustomMessage(type = 'refresh_file_list', message = list())
  })

  # ── Dependency Status ──────────────────────────────────────────────────────
  output$dependency_status <- renderUI({
    py_status <- python_env_status
    
    # Check R packages
    r_check <- all(sapply(REQUIRED_R_PACKAGES, requireNamespace, quietly = TRUE))
    
    html_content <- ""
    
    # R Packages Status
    html_content <- paste0(html_content,
      "<strong style='color:#2c3e50;'>R-Pakete:</strong><br/>",
      if (r_check) {
        "<span style='color:#27ae60;'>✓ Alle erforderlich Pakete installiert</span>"
      } else {
        "<span style='color:#e74c3c;'>✗ Einige R-Pakete fehlen</span>"
      }, "<br/><br/>"
    )
    
    # Python Status
    if (py_status$python_available) {
      html_content <- paste0(html_content,
        "<strong style='color:#2c3e50;'>Python:</strong><br/>",
        "<span style='color:#27ae60;'>✓ Python verfügbar: ", py_status$python_version, "</span><br/>"
      )
      
      if (py_status$all_packages_available) {
        html_content <- paste0(html_content,
          "<span style='color:#27ae60;'>✓ Alle Python-Pakete installiert:</span><br/>",
          "   • pdfplumber<br/>",
          "   • openpyxl<br/>"
        )
      } else {
        html_content <- paste0(html_content,
          "<span style='color:#f39c12;'>⚠ Einige Python-Pakete fehlen:</span><br/>",
          paste0("   • ", py_status$missing_packages, "<br/>", collapse = "")
        )
      }
    } else {
      html_content <- paste0(html_content,
        "<strong style='color:#2c3e50;'>Python:</strong><br/>",
        "<span style='color:#e74c3c;'>✗ Python 3 nicht gefunden. Bitte installieren Sie Python 3.</span><br/>"
      )
    }
    
    # Optional pdftools
    html_content <- paste0(html_content,
      "<br/><strong style='color:#2c3e50;'>Optional (PDF-Vergleich):</strong><br/>",
      if (HAS_PDFTOOLS) {
        "<span style='color:#27ae60;'>✓ pdftools und png installiert</span>"
      } else {
        "<span style='color:#f39c12;'>⚠ pdftools/png nicht installiert (PDF-Vergleich deaktiviert)</span>"
      }
    )
    
    HTML(html_content)
  })

  # ── PDF Upload and Extraction ──────────────────────────────────────────────
  extract_status_msg <- reactiveVal("Bereit für PDF-Upload")
  extraction_summary_msg <- reactiveVal("Keine Extraktion durchgeführt")
  
  output$extract_status <- renderText({
    extract_status_msg()
  })
  
  output$extraction_summary <- renderText({
    extraction_summary_msg()
  })
  
  observeEvent(input$btn_extract, {
    req(input$pdf_upload)
    
    # Check Python availability
    if (!python_env_status$python_available) {
      extract_status_msg("Fehler: Python 3 nicht verfügbar")
      return()
    }
    
    # Install missing Python packages if needed
    if (!python_env_status$all_packages_available) {
      extract_status_msg("Installiere fehlende Python-Pakete...")
      tryCatch({
        install_python_packages(python_env_status$missing_packages)
        extract_status_msg("Python-Pakete installiert. Starte Extraktion...")
      }, error = function(e) {
        extract_status_msg(paste("Fehler beim Installieren von Python-Paketen:", e$message))
        return()
      })
    }
    
    # Create temp directory for uploads
    upload_dir <- tempfile(pattern = "pdf_upload_")
    dir.create(upload_dir, recursive = TRUE)
    
    # Copy uploaded files to temp directory
    pdf_files <- character(0)
    for (i in seq_len(nrow(input$pdf_upload))) {
      src <- input$pdf_upload$datapath[i]
      dst <- file.path(upload_dir, input$pdf_upload$name[i])
      file.copy(src, dst)
      pdf_files <- c(pdf_files, dst)
    }
    
    # Extract tables from each PDF with progress bar
    total_tables <- 0
    errors <- character(0)
    n_pdfs <- length(pdf_files)
    
    # Use withProgress for visual feedback
    withProgress(
      message = "📊 Tabellen extrahieren",
      detail = sprintf("0 / %d PDFs verarbeitet", n_pdfs),
      value = 0,
      {
        for (idx in seq_along(pdf_files)) {
          pdf_file <- pdf_files[idx]
          pdf_name <- basename(pdf_file)
          
          # Update progress message
          incProgress(
            amount = 0,
            detail = sprintf(
              "(%d / %d) Datei: %s",
              idx,
              n_pdfs,
              pdf_name
            )
          )
          
          # Run Python extraction script
          tryCatch({
            cmd_output <- system2(
              "python3",
              c("extract_tables.py", pdf_file, "extracted_tables"),
              stdout = TRUE,
              stderr = TRUE,
              cwd = getwd()
            )
            
            # Count extracted tables from output
            table_lines <- grep("→", cmd_output, value = TRUE)
            num_tables <- length(table_lines)
            total_tables <- total_tables + num_tables
            
            # Update status message with extraction details
            extract_status_msg(
              sprintf(
                "✓ %s\n  Tabellen: %d | Gesamt: %d",
                pdf_name,
                num_tables,
                total_tables
              )
            )
            
          }, error = function(e) {
            errors <<- c(errors, sprintf("%s: %s", pdf_name, e$message))
            extract_status_msg(
              sprintf(
                "✗ %s\n  Fehler: %s",
                pdf_name,
                e$message
              )
            )
          })
          
          # Increment progress bar
          incProgress(
            amount = 1 / n_pdfs,
            detail = sprintf(
              "(%d / %d) %d Tabelle(n) gefunden",
              idx,
              n_pdfs,
              total_tables
            )
          )
        }
      }
    )
    
    # Update UI with final results
    if (length(errors) > 0L) {
      summary_text <- sprintf(
        "⚠️ Extraktion abgeschlossen mit Fehlern.\n\nInsgesamt %d Tabelle(n) extrahiert.\n\n🔴 Fehler:\n%s",
        total_tables,
        paste(errors, collapse = "\n")
      )
      extract_status_msg(sprintf("⚠️ %d Tabelle(n) extrahiert mit %d Fehler(n)", 
                                 total_tables, length(errors)))
    } else {
      summary_text <- sprintf(
        "✅ Extraktion erfolgreich abgeschlossen!\n\n🎉 Insgesamt %d Tabelle(n) extrahiert.\n\n📁 Ort: extracted_tables/\n\n→ Jetzt zur 'Tabellen & Vorschau' Tab wechseln um die Tabellen zu verarbeiten.",
        total_tables
      )
      extract_status_msg(sprintf("✅ %d Tabelle(n) erfolgreich extrahiert", total_tables))
    }
    
    extraction_summary_msg(summary_text)
    
    # Refresh file list
    session$sendCustomMessage(type = 'refresh_file_list', message = list())
  })

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

  # ── Formatting Rules (Numeric) ────────────────────────────────────────────
  fmt_rules_rv <- reactiveVal(rules_to_df(load_fmt_rules()))

  output$fmt_rules_tbl <- renderDT({
    datatable(
      fmt_rules_rv(),
      editable  = "cell",
      rownames  = FALSE,
      selection = "multiple",
      options   = list(pageLength = 15, scrollX = TRUE),
      colnames  = c("Name", "Pattern (Regex)", "Replacement", "Description", "Enabled"),
      class     = "compact stripe"
    )
  })

  observeEvent(input$fmt_rules_tbl_cell_edit, {
    info <- input$fmt_rules_tbl_cell_edit
    df   <- fmt_rules_rv()
    value <- info$value
    # Convert "on"/"off" strings to TRUE/FALSE for the enabled column
    if (info$col + 1L == 5L) {
      value <- tolower(value) %in% c("true", "1", "on", "yes")
    }
    df[info$row, info$col + 1L] <- value
    fmt_rules_rv(df)
  })

  observeEvent(input$btn_save_fmt_rules, {
    save_fmt_rules(fmt_rules_rv())
    showNotification("Formatierungsregeln in formatting_rules_de.csv gespeichert.",
                     type = "message", duration = 3)
  })

  observeEvent(input$btn_reset_fmt_rules, {
    fmt_rules_rv(rules_to_df(FORMATTING_RULES_DE))
    showNotification("Standard-Formatierungsregeln wiederhergestellt.",
                     type = "warning", duration = 3)
  })

  observeEvent(input$btn_add_fmt_rule, {
    df <- fmt_rules_rv()
    df <- rbind(df, data.frame(
      name = "New Rule",
      pattern = "pattern",
      replacement = "replacement",
      description = "Description",
      enabled = TRUE,
      stringsAsFactors = FALSE
    ))
    fmt_rules_rv(df)
  })

  observeEvent(input$btn_del_fmt_rules, {
    sel <- input$fmt_rules_tbl_rows_selected
    if (!is.null(sel) && length(sel) > 0L) {
      df <- fmt_rules_rv()
      fmt_rules_rv(df[-sel, , drop = FALSE])
    }
  })

  # Convert formatting rules dataframe back to list format for use in build_flextable
  fmt_rules_list_rv <- reactive({ df_to_rules(fmt_rules_rv()) })

  # ── Custom Styles Management ──────────────────────────────────────────────
  custom_styles_rv <- reactiveVal(load_custom_styles())
  custom_styles_invalidate <- reactiveVal(0)

  # Dynamic style selector for main panel
  output$style_selector_ui <- renderUI({
    styles <- get_available_styles()
    selectInput("style_id", label = NULL, choices = styles, selected = "clinical")
  })

  # Dynamic style selector for batch export
  output$batch_style_selector_ui <- renderUI({
    styles <- get_available_styles()
    selectInput("b_style", label = NULL, choices = styles, selected = "clinical")
  })

  # Display custom styles table
  output$custom_styles_tbl <- renderDT({
    custom_styles_invalidate()  # Add dependency
    df <- custom_styles_rv()
    if (nrow(df) == 0L) {
      return(datatable(
        get_empty_custom_styles_df(),
        options = list(dom = "t"),
        colnames = c("Name", "Schriftart", "Größe", "Kopf. Fett", "Kopf. Farbe", 
                     "Kopf. Hintergrund", "Datenzellen Hintergrund", "Wechselnde", "Alternativ", "Padding"),
        class = "compact stripe"
      ))
    }
    datatable(
      df,
      selection = "single",
      rownames = FALSE,
      options = list(pageLength = 10, scrollX = TRUE),
      colnames = c("Name", "Schriftart", "Größe", "Kopf. Fett", "Kopf. Farbe", 
                   "Kopf. Hintergrund", "Datenzellen Hintergrund", "Wechselnde", "Alternativ", "Padding"),
      class = "compact stripe"
    )
  })

  # Load selected style for editing
  observeEvent(input$custom_styles_tbl_rows_selected, {
    sel <- input$custom_styles_tbl_rows_selected
    if (is.null(sel) || length(sel) == 0L) return()
    
    df <- custom_styles_rv()
    if (sel > nrow(df)) return()
    
    row <- df[sel, ]
    updateTextInput(session, "custom_style_name", value = row$name)
    updateSelectInput(session, "cs_font_name", selected = row$font_name)
    updateNumericInput(session, "cs_font_size", value = row$font_size)
    updateCheckboxInput(session, "cs_header_bold", value = row$header_bold)
    colourpicker::updateColourInput(session, "cs_header_color", value = row$header_color)
    colourpicker::updateColourInput(session, "cs_header_bg", value = row$header_bg)
    colourpicker::updateColourInput(session, "cs_body_bg", value = row$body_bg)
    updateCheckboxInput(session, "cs_alternating", value = row$alternating_rows)
    colourpicker::updateColourInput(session, "cs_alternating_bg", value = row$alternating_bg)
    updateNumericInput(session, "cs_padding", value = row$padding)
  })

  # Save custom style
  observeEvent(input$btn_save_custom_style, {
    name <- trimws(input$custom_style_name)
    if (!nzchar(name)) {
      showNotification("Bitte einen Stilnamen eingeben.", type = "warning", duration = 3)
      return()
    }
    
    new_style <- data.frame(
      name = name,
      font_name = input$cs_font_name,
      font_size = input$cs_font_size,
      header_bold = input$cs_header_bold,
      header_color = input$cs_header_color,
      header_bg = input$cs_header_bg,
      body_bg = input$cs_body_bg,
      alternating_rows = input$cs_alternating,
      alternating_bg = input$cs_alternating_bg,
      padding = input$cs_padding,
      stringsAsFactors = FALSE
    )
    
    df <- custom_styles_rv()
    # Check if updating existing style
    existing_idx <- which(df$name == name)
    if (length(existing_idx) > 0L) {
      df[existing_idx[1L], ] <- new_style
    } else {
      df <- rbind(df, new_style)
    }
    
    tryCatch({
      save_custom_styles(df)
      custom_styles_rv(df)
      custom_styles_invalidate(custom_styles_invalidate() + 1)
      showNotification(sprintf("Stil '%s' gespeichert.", name), type = "message", duration = 3)
    }, error = function(e) {
      showNotification(sprintf("Fehler beim Speichern: %s", e$message), type = "error", duration = 4)
    })
  })

  # Delete custom style
  observeEvent(input$btn_delete_custom_style, {
    sel <- input$custom_styles_tbl_rows_selected
    if (is.null(sel) || length(sel) == 0L) {
      showNotification("Bitte einen Stil auswählen.", type = "warning", duration = 3)
      return()
    }
    
    df <- custom_styles_rv()
    style_name <- df$name[sel]
    
    tryCatch({
      if (delete_custom_style(style_name)) {
        df <- load_custom_styles()
        custom_styles_rv(df)
        custom_styles_invalidate(custom_styles_invalidate() + 1)
        showNotification(sprintf("Stil '%s' gelöscht.", style_name), type = "message", duration = 3)
      }
    }, error = function(e) {
      showNotification(sprintf("Fehler beim Löschen: %s", e$message), type = "error", duration = 4)
    })
  })

  # Preview custom style
  output$custom_style_preview <- renderUI({
    # Create a minimal preview table
    preview_df <- data.frame(
      Parameter = c("Wert 1", "Wert 2", "Wert 3"),
      Ergebnis = c("12.5%", "25.3%", "40.8%"),
      check.names = FALSE
    )
    
    ft <- flextable::flextable(preview_df)
    
    # Apply custom style preview properties
    font_nm <- if (is.null(input$cs_font_name)) "Arial" else input$cs_font_name
    font_sz <- if (is.null(input$cs_font_size)) 9 else input$cs_font_size
    pad     <- if (is.null(input$cs_padding)) 3 else input$cs_padding
    pad_lr  <- pad + 1L
    
    ft <- flextable::font(ft, fontname = font_nm, part = "all")
    ft <- flextable::fontsize(ft, size = font_sz, part = "all")
    ft <- flextable::align(ft, align = "left", part = "all")
    ft <- flextable::padding(ft, padding.top = pad, padding.bottom = pad,
                             padding.left = pad_lr, padding.right = pad_lr, part = "all")
    ft <- flextable::border_remove(ft)
    
    nc <- length(ft$col_keys)
    if (nc >= 2L) {
      data_j <- seq(2L, nc)
      ft <- flextable::align(ft, j = data_j, align = "right", part = "body")
      ft <- flextable::align(ft, j = data_j, align = "center", part = "header")
    }
    
    # Header styling
    header_bold <- if (is.null(input$cs_header_bold)) TRUE else input$cs_header_bold
    header_color <- if (is.null(input$cs_header_color)) "#ffffff" else input$cs_header_color
    header_bg <- if (is.null(input$cs_header_bg)) "#1a3a5c" else input$cs_header_bg
    body_bg <- if (is.null(input$cs_body_bg)) "#ffffff" else input$cs_body_bg
    alternating <- if (is.null(input$cs_alternating)) FALSE else input$cs_alternating
    alternating_bg <- if (is.null(input$cs_alternating_bg)) "#f0f5fb" else input$cs_alternating_bg
    
    if (header_bold) {
      ft <- flextable::bold(ft, bold = TRUE, part = "header")
    }
    ft <- flextable::color(ft, color = header_color, part = "header")
    ft <- flextable::bg(ft, bg = header_bg, part = "header")
    
    # Body styling
    ft <- flextable::bg(ft, bg = body_bg, part = "body")
    
    # Alternating rows
    if (alternating) {
      nr_b <- flextable::nrow_part(ft, "body")
      if (nr_b > 0L) {
        odd  <- seq(1L, nr_b, 2L)
        even <- seq(2L, nr_b, 2L)
        if (length(odd))  ft <- flextable::bg(ft, i = odd, bg = alternating_bg, part = "body")
        if (length(even)) ft <- flextable::bg(ft, i = even, bg = body_bg, part = "body")
      }
    }
    
    ft <- flextable::set_table_properties(ft, layout = "autofit", width = 1)
    
    # Apply borders
    b_thick <- fp_border(color = "#111111", width = 1.5)
    ft <- flextable::border_remove(ft)
    ft <- flextable::hline_top(ft, border = b_thick, part = "header")
    ft <- flextable::hline_bottom(ft, border = b_thick, part = "header")
    ft <- flextable::hline_bottom(ft, border = b_thick, part = "body")
    
    htmltools_value(ft)
  })

  # ── Custom Styles Management ──────────────────────────────────────────────
  custom_styles_rv <- reactiveVal(load_custom_styles())
  custom_styles_invalidate <- reactiveVal(0)

  # Dynamic style selector for main panel
  output$style_selector_ui <- renderUI({
    styles <- get_available_styles()
    selectInput("style_id", label = NULL, choices = styles, selected = "clinical")
  })

  # Dynamic style selector for batch export
  output$batch_style_selector_ui <- renderUI({
    styles <- get_available_styles()
    selectInput("b_style", label = NULL, choices = styles, selected = "clinical")
  })

  # Display custom styles table
  output$custom_styles_tbl <- renderDT({
    custom_styles_invalidate()  # Add dependency
    df <- custom_styles_rv()
    if (nrow(df) == 0L) {
      return(datatable(
        get_empty_custom_styles_df(),
        options = list(dom = "t"),
        colnames = c("Name", "Schriftart", "Größe", "Kopf. Fett", "Kopf. Farbe", 
                     "Kopf. Hintergrund", "Datenzellen Hintergrund", "Wechselnde", "Alternativ", "Padding"),
        class = "compact stripe"
      ))
    }
    datatable(
      df,
      selection = "single",
      rownames = FALSE,
      options = list(pageLength = 10, scrollX = TRUE),
      colnames = c("Name", "Schriftart", "Größe", "Kopf. Fett", "Kopf. Farbe", 
                   "Kopf. Hintergrund", "Datenzellen Hintergrund", "Wechselnde", "Alternativ", "Padding"),
      class = "compact stripe"
    )
  })

  # Load selected style for editing
  observeEvent(input$custom_styles_tbl_rows_selected, {
    sel <- input$custom_styles_tbl_rows_selected
    if (is.null(sel) || length(sel) == 0L) return()
    
    df <- custom_styles_rv()
    if (sel > nrow(df)) return()
    
    row <- df[sel, ]
    updateTextInput(session, "custom_style_name", value = row$name)
    updateSelectInput(session, "cs_font_name", selected = row$font_name)
    updateNumericInput(session, "cs_font_size", value = row$font_size)
    updateCheckboxInput(session, "cs_header_bold", value = row$header_bold)
    colourpicker::updateColourInput(session, "cs_header_color", value = row$header_color)
    colourpicker::updateColourInput(session, "cs_header_bg", value = row$header_bg)
    colourpicker::updateColourInput(session, "cs_body_bg", value = row$body_bg)
    updateCheckboxInput(session, "cs_alternating", value = row$alternating_rows)
    colourpicker::updateColourInput(session, "cs_alternating_bg", value = row$alternating_bg)
    updateNumericInput(session, "cs_padding", value = row$padding)
  })

  # Save custom style
  observeEvent(input$btn_save_custom_style, {
    name <- trimws(input$custom_style_name)
    if (!nzchar(name)) {
      showNotification("Bitte einen Stilnamen eingeben.", type = "warning", duration = 3)
      return()
    }
    
    new_style <- data.frame(
      name = name,
      font_name = input$cs_font_name,
      font_size = input$cs_font_size,
      header_bold = input$cs_header_bold,
      header_color = input$cs_header_color,
      header_bg = input$cs_header_bg,
      body_bg = input$cs_body_bg,
      alternating_rows = input$cs_alternating,
      alternating_bg = input$cs_alternating_bg,
      padding = input$cs_padding,
      stringsAsFactors = FALSE
    )
    
    df <- custom_styles_rv()
    # Check if updating existing style
    existing_idx <- which(df$name == name)
    if (length(existing_idx) > 0L) {
      df[existing_idx[1L], ] <- new_style
    } else {
      df <- rbind(df, new_style)
    }
    
    tryCatch({
      save_custom_styles(df)
      custom_styles_rv(df)
      custom_styles_invalidate(custom_styles_invalidate() + 1)
      showNotification(sprintf("Stil '%s' gespeichert.", name), type = "message", duration = 3)
    }, error = function(e) {
      showNotification(sprintf("Fehler beim Speichern: %s", e$message), type = "error", duration = 4)
    })
  })

  # Delete custom style
  observeEvent(input$btn_delete_custom_style, {
    sel <- input$custom_styles_tbl_rows_selected
    if (is.null(sel) || length(sel) == 0L) {
      showNotification("Bitte einen Stil auswählen.", type = "warning", duration = 3)
      return()
    }
    
    df <- custom_styles_rv()
    style_name <- df$name[sel]
    
    tryCatch({
      if (delete_custom_style(style_name)) {
        df <- load_custom_styles()
        custom_styles_rv(df)
        custom_styles_invalidate(custom_styles_invalidate() + 1)
        showNotification(sprintf("Stil '%s' gelöscht.", style_name), type = "message", duration = 3)
      }
    }, error = function(e) {
      showNotification(sprintf("Fehler beim Löschen: %s", e$message), type = "error", duration = 4)
    })
  })

  # Preview custom style
  output$custom_style_preview <- renderUI({
    # Create a minimal preview table
    preview_df <- data.frame(
      Parameter = c("Wert 1", "Wert 2", "Wert 3"),
      Ergebnis = c("12.5%", "25.3%", "40.8%"),
      check.names = FALSE
    )
    
    ft <- flextable::flextable(preview_df)
    
    # Apply custom style preview properties
    font_nm <- if (is.null(input$cs_font_name)) "Arial" else input$cs_font_name
    font_sz <- if (is.null(input$cs_font_size)) 9 else input$cs_font_size
    pad     <- if (is.null(input$cs_padding)) 3 else input$cs_padding
    pad_lr  <- pad + 1L
    
    ft <- flextable::font(ft, fontname = font_nm, part = "all")
    ft <- flextable::fontsize(ft, size = font_sz, part = "all")
    ft <- flextable::align(ft, align = "left", part = "all")
    ft <- flextable::padding(ft, padding.top = pad, padding.bottom = pad,
                             padding.left = pad_lr, padding.right = pad_lr, part = "all")
    ft <- flextable::border_remove(ft)
    
    nc <- length(ft$col_keys)
    if (nc >= 2L) {
      data_j <- seq(2L, nc)
      ft <- flextable::align(ft, j = data_j, align = "right", part = "body")
      ft <- flextable::align(ft, j = data_j, align = "center", part = "header")
    }
    
    # Header styling
    header_bold <- if (is.null(input$cs_header_bold)) TRUE else input$cs_header_bold
    header_color <- if (is.null(input$cs_header_color)) "#ffffff" else input$cs_header_color
    header_bg <- if (is.null(input$cs_header_bg)) "#1a3a5c" else input$cs_header_bg
    body_bg <- if (is.null(input$cs_body_bg)) "#ffffff" else input$cs_body_bg
    alternating <- if (is.null(input$cs_alternating)) FALSE else input$cs_alternating
    alternating_bg <- if (is.null(input$cs_alternating_bg)) "#f0f5fb" else input$cs_alternating_bg
    
    if (header_bold) {
      ft <- flextable::bold(ft, bold = TRUE, part = "header")
    }
    ft <- flextable::color(ft, color = header_color, part = "header")
    ft <- flextable::bg(ft, bg = header_bg, part = "header")
    
    # Body styling
    ft <- flextable::bg(ft, bg = body_bg, part = "body")
    
    # Alternating rows
    if (alternating) {
      nr_b <- flextable::nrow_part(ft, "body")
      if (nr_b > 0L) {
        odd  <- seq(1L, nr_b, 2L)
        even <- seq(2L, nr_b, 2L)
        if (length(odd))  ft <- flextable::bg(ft, i = odd, bg = alternating_bg, part = "body")
        if (length(even)) ft <- flextable::bg(ft, i = even, bg = body_bg, part = "body")
      }
    }
    
    ft <- flextable::set_table_properties(ft, layout = "autofit", width = 1)
    
    # Apply borders
    b_thick <- fp_border(color = "#111111", width = 1.5)
    ft <- flextable::border_remove(ft)
    ft <- flextable::hline_top(ft, border = b_thick, part = "header")
    ft <- flextable::hline_bottom(ft, border = b_thick, part = "header")
    ft <- flextable::hline_bottom(ft, border = b_thick, part = "body")
    
    htmltools_value(ft)
  })

  # ── Pattern (Muster) Management ───────────────────────────────────────────
  patterns_invalidate <- reactiveVal(0)  # Trigger for pattern list refresh
  patterns_list_rv <- reactive({
    patterns_invalidate()  # Dependency for reactivity
    list_patterns()
  })

  output$pattern_list_ui <- renderUI({
    patterns <- patterns_list_rv()
    if (length(patterns) == 0L) {
      return(div(style = "color:#999; padding:10px; background:#f5f5f5; border-radius:4px;",
                 "Noch keine Muster vorhanden"))
    }
    selectInput("pattern_select", label = NULL,
                choices = setNames(patterns, patterns))
  })

  observeEvent(input$btn_save_pattern, {
    name <- trimws(input$pattern_save_name)
    if (!nzchar(name)) {
      showNotification("Bitte einen Namen für das Muster eingeben.",
                       type = "warning", duration = 3)
      return()
    }
    
    tryCatch({
      save_pattern(name, trans_rv(), fmt_rules_rv(), input$style_id, custom_styles_rv())
      save_pattern(name, trans_rv(), fmt_rules_rv(), input$style_id, custom_styles_rv())
      showNotification(sprintf("Muster '%s' gespeichert.", name),
                       type = "message", duration = 3)
      updateTextInput(session, "pattern_save_name", value = "")
      patterns_invalidate(patterns_invalidate() + 1)  # Trigger refresh
    }, error = function(e) {
      showNotification(sprintf("Fehler beim Speichern: %s", e$message),
                       type = "error", duration = 4)
    })
  })

  observeEvent(input$btn_load_pattern, {
    selected <- input$pattern_select
    if (is.null(selected) || !nzchar(selected)) {
      showNotification("Bitte ein Muster auswählen.",
                       type = "warning", duration = 3)
      return()
    }
    
    tryCatch({
      pattern_data <- load_pattern(selected)
      if (is.null(pattern_data)) {
        showNotification("Muster konnte nicht geladen werden.",
                         type = "error", duration = 3)
        return()
      }
      
      # Load style
      if (!is.null(pattern_data$style)) {
        updateSelectInput(session, "style_id", selected = pattern_data$style)
      }
      
      # Load translations
      if (!is.null(pattern_data$translations)) {
        trans_rv(pattern_data$translations)
      }
      
      # Load formatting rules
      if (!is.null(pattern_data$fmt_rules)) {
        fmt_rules_rv(pattern_data$fmt_rules)
      }
      
      # Load custom styles
      if (!is.null(pattern_data$custom_styles)) {
        custom_styles_rv(pattern_data$custom_styles)
        custom_styles_invalidate(custom_styles_invalidate() + 1)
      }
      
      # Load custom styles
      if (!is.null(pattern_data$custom_styles)) {
        custom_styles_rv(pattern_data$custom_styles)
        custom_styles_invalidate(custom_styles_invalidate() + 1)
      }
      
      showNotification(sprintf("Muster '%s' geladen.", selected),
                       type = "message", duration = 3)
    }, error = function(e) {
      showNotification(sprintf("Fehler beim Laden: %s", e$message),
                       type = "error", duration = 4)
    })
  })

  observeEvent(input$btn_delete_pattern, {
    selected <- input$pattern_select
    if (is.null(selected) || !nzchar(selected)) {
      showNotification("Bitte ein Muster auswählen.",
                       type = "warning", duration = 3)
      return()
    }
    
    showModal(modalDialog(
      title = "Muster löschen",
      sprintf("Wirklich Muster '%s' löschen?", selected),
      footer = tagList(
        actionButton("confirm_delete_pattern", "Ja, löschen",
                     class = "btn-danger"),
        modalButton("Abbrechen", class = "btn-secondary")
      )
    ))
  })

  observeEvent(input$confirm_delete_pattern, {
    selected <- input$pattern_select
    tryCatch({
      if (delete_pattern(selected)) {
        showNotification(sprintf("Muster '%s' gelöscht.", selected),
                         type = "message", duration = 3)
        removeModal()
        patterns_invalidate(patterns_invalidate() + 1)  # Trigger refresh
        # Clear the selection after deleting
        updateSelectInput(session, "pattern_select", selected = "")
      } else {
        showNotification("Muster konnte nicht gelöscht werden.",
                         type = "error", duration = 3)
      }
    }, error = function(e) {
      showNotification(sprintf("Fehler beim Löschen: %s", e$message),
                       type = "error", duration = 4)
    })
  })

  output$pattern_info_text <- renderText({
    patterns <- patterns_list_rv()
    selected <- input$pattern_select
    
    if (length(patterns) == 0L) {
      return("Keine Muster vorhanden.\nErstelle dein erstes Muster mit\ndem Namen und klicke 'Muster speichern'.")
    }
    
    if (is.null(selected) || !nzchar(selected)) {
      sprintf("Gespeicherte Muster: %d\n\nWähle ein Muster oben aus\num Details zu sehen.",
              length(patterns))
    } else {
      pattern_path <- file.path(PATTERNS_DIR, selected)
      file_info <- file.info(pattern_path)
      mod_time <- if (!is.na(file_info$mtime)) {
        format(file_info$mtime, "%d.%m.%Y %H:%M")
      } else {
        "Unbekannt"
      }
      
      sprintf("Muster: %s\n\nGeändert: %s\nPfad: %s",
              selected, mod_time, pattern_path)
    }
  })

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
    fmt_rules <- if (input$language == "DE") fmt_rules_list_rv() else NULL
    ft  <- tryCatch(
      build_flextable(tbl, input$style_id, input$language, tmap_rv(), fmt_rules),
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
      fmt_rules <- if (input$language == "DE") fmt_rules_list_rv() else NULL
      ft  <- build_flextable(tbl, input$style_id, input$language, tmap_rv(), fmt_rules)
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

      fmt_rules <- if (bln == "DE") fmt_rules_list_rv() else NULL
      withProgress(message = "Erzeuge RTF-Dateien …", value = 0, {
        for (idx in seq_along(rows)) {
          r  <- rows[idx]
          fp <- file.path(EXCEL_DIR, df$Datei[r])
          tryCatch({
            tbl <- read_excel_table(fp)
            ft  <- build_flextable(tbl, bst, bln, tmap, fmt_rules)
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

  # ── Combine RTFs with TOC and Bookmarks ────────────────────────────────────
  # Table for combine selection
  output$comb_tbl <- renderDT({
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

  # Select/deselect all for combine
  observeEvent(input$comb_all, {
    proxy <- dataTableProxy("comb_tbl")
    selectRows(proxy, seq_len(nrow(file_df())))
  })

  observeEvent(input$comb_none, {
    proxy <- dataTableProxy("comb_tbl")
    selectRows(proxy, NULL)
  })

  # Selection info for combine
  output$comb_sel_info <- renderUI({
    n <- length(input$comb_tbl_rows_selected)
    div(class = "info-box",
        sprintf("%d Tabelle(n) ausgewählt", n))
  })

  # Function to create table of contents with bookmarks
  create_combined_rtf_with_toc <- function(file, rows, df, tmap, fmt_rules, 
                                           title, include_toc, include_bookmarks, page_breaks) {
    if (is.null(rows) || length(rows) == 0L) {
      showNotification("Keine Tabellen ausgewählt.", type = "warning")
      return(FALSE)
    }

    tryCatch({
      # Collect all tables first for TOC
      tables_data <- list()
      toc_entries <- character(0)
      
      for (idx in seq_along(rows)) {
        r  <- rows[idx]
        fp <- file.path(EXCEL_DIR, df$Datei[r])
        tbl <- read_excel_table(fp)
        ft <- build_flextable(tbl, "clinical", "DE", tmap, fmt_rules)
        
        tables_data[[idx]] <- list(
          title = tbl$title,
          flextable = ft,
          extra_hdrs = tbl$extra_hdrs,
          footnotes = tbl$footnotes
        )
        
        toc_entries <- c(toc_entries, tbl$title)
      }
      
      # Start building document
      doc <- officer::read_docx()
      
      # Add title
      title_run <- officer::run_pagebreak()
      doc <- officer::body_add_par(doc, title, 
                                   style = "Heading 1",
                                   run_properties = officer::fp_text(
                                     bold = TRUE, 
                                     size = 24, 
                                     color = "#1a3a5c"))
      
      # Add date
      doc <- officer::body_add_par(doc, 
                                  paste("Erstellt am:", format(Sys.Date(), "%d.%m.%Y")),
                                  style = "Normal")
      doc <- officer::body_add_par(doc, "")
      
      # Add TOC if requested
      if (include_toc) {
        doc <- officer::body_add_break(doc)
        
        doc <- officer::body_add_par(doc, "Inhaltsverzeichnis", 
                                     style = "Heading 2",
                                     run_properties = officer::fp_text(
                                       bold = TRUE, 
                                       size = 14,
                                       color = "#2c3e50"))
        doc <- officer::body_add_par(doc, "")
        
        # Add TOC entries
        for (i in seq_along(toc_entries)) {
          entry_text <- paste0(i, ". ", toc_entries[i])
          doc <- officer::body_add_par(doc, entry_text,
                                      style = "Normal")
        }
        
        doc <- officer::body_add_break(doc)
      }
      
      # Add tables
      for (idx in seq_along(tables_data)) {
        tbl_data <- tables_data[[idx]]
        
        # Add table title/heading with bookmark indicator
        if (include_bookmarks) {
          title_with_bookmark <- paste0("[", idx, "] ", tbl_data$title)
        } else {
          title_with_bookmark <- paste0("[", idx, "] ", tbl_data$title)
        }
        
        doc <- officer::body_add_par(doc, title_with_bookmark, 
                                     style = "Heading 3",
                                     run_properties = officer::fp_text(
                                       bold = TRUE, 
                                       size = 12,
                                       color = "#2c3e50"))
        
        # Add table
        doc <- officer::body_add_flextable(doc, tbl_data$flextable)
        
        # Add extra headers if any
        if (length(tbl_data$extra_hdrs) > 0L) {
          for (eh in tbl_data$extra_hdrs) {
            doc <- officer::body_add_par(doc, paste("*", eh), 
                                        style = "Normal",
                                        run_properties = officer::fp_text(
                                          size = 9, 
                                          italic = TRUE, 
                                          color = "#666"))
          }
        }
        
        # Add footnotes if any
        if (length(tbl_data$footnotes) > 0L) {
          doc <- officer::body_add_par(doc, "")
          for (fn in tbl_data$footnotes) {
            doc <- officer::body_add_par(doc, fn, 
                                        style = "Normal",
                                        run_properties = officer::fp_text(
                                          size = 9, 
                                          italic = TRUE, 
                                          color = "#777"))
          }
        }
        
        # Add spacing or page break
        if (idx < length(tables_data)) {
          if (page_breaks) {
            doc <- officer::body_add_break(doc)
          } else {
            doc <- officer::body_add_par(doc, "")
            doc <- officer::body_add_par(doc, "")
          }
        }
      }
      
      # Save as DOCX
      officer::print(doc, target = file)
      
      return(TRUE)
    }, error = function(e) {
      showNotification(paste("Fehler beim Kombinieren:", e$message), 
                      type = "error", duration = 5)
      return(FALSE)
    })
  }

  # Download handler for combined RTF
  output$dl_combined_rtf <- downloadHandler(
    filename = function() {
      paste0("Kombiniert_", input$combined_rtf_title, "_",
             format(Sys.Date(), "%Y%m%d"), ".docx")
    },
    content = function(file) {
      withProgress(message = "Kombiniere RTF-Dateien …", value = 0, {
        rows <- input$comb_tbl_rows_selected
        if (is.null(rows) || length(rows) == 0L) {
          showNotification("Keine Tabellen ausgewählt.", type = "warning")
          writeLines("No tables selected.", file)
          return()
        }

        df <- file_df()
        tmap <- tmap_rv()
        fmt_rules <- fmt_rules_list_rv()
        
        create_combined_rtf_with_toc(
          file = file,
          rows = rows,
          df = df,
          tmap = tmap,
          fmt_rules = fmt_rules,
          title = input$combined_rtf_title,
          include_toc = input$comb_include_toc,
          include_bookmarks = input$comb_include_bookmarks,
          page_breaks = input$comb_page_breaks
        )
        
        showNotification("✅ Tabellen erfolgreich kombiniert!", type = "message", duration = 3)
      })
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
        column(7,
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
              imageOutput("cmp_pdf_img", height = "100%", width = "100%"))
        ),
        column(5,
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
    fmt_rules <- if (input$language == "DE") fmt_rules_list_rv() else NULL
    ft  <- tryCatch(
      build_flextable(tbl, input$style_id, input$language, tmap_rv(), fmt_rules),
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
