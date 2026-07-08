Vibe coding of a Shinyapp that extracts interaction p values from PDF files and displays them per endpoint

first run:
$ python3 extract_tables.pdf [path to pdf] [output path]
then run:
$ R -e "shiny::runApp('interaction_pvalues_app.R', launch.browser=TRUE)"
