library(shiny)
library(shinydashboard)
library(ggplot2)
library(dplyr)
library(sf)
library(ggpubr)
library(survival)
library(survminer)
library(flextable)
library(officer)
library(DT)
library(haven)
library(readxl)
library(plotly)
library(colourpicker)
library(epitools)
library(tidyverse)
library(broom)
library(car)
library(emmeans)

# UI
ui <- dashboardPage(
  dashboardHeader(title = "📈 EpiDem Suite™"),
  dashboardSidebar(
    sidebarMenu(
      menuItem("Data Upload", tabName = "data", icon = icon("database")),
      menuItem("Descriptive Analyses", tabName = "descriptive", icon = icon("chart-bar"),
               menuSubItem("Frequencies", tabName = "freq"),
               menuSubItem("Central Tendency", tabName = "central"),
               menuSubItem("Epidemic Curves", tabName = "epicurve"),
               menuSubItem("Spatial Maps", tabName = "maps"),
               menuSubItem("Demographics", tabName = "demo")),
      menuItem("Disease Frequency", tabName = "disease", icon = icon("viruses"),
               menuSubItem("Incidence/Prevalence", tabName = "incprev"),
               menuSubItem("Attack Rates", tabName = "attack"),
               menuSubItem("Mortality Rates", tabName = "mortality"),
               menuSubItem("Life Table Analysis", tabName = "lifetable")),
      menuItem("Measures of Association", tabName = "association", icon = icon("balance-scale"),
               menuSubItem("Risk Ratios", tabName = "risk"),
               menuSubItem("Odds Ratios", tabName = "odds"),
               menuSubItem("Attributable Risk", tabName = "ar"),
               menuSubItem("Poisson Regression", tabName = "poisson")),
      menuItem("Survival Analysis", tabName = "survival", icon = icon("clock"),
               menuSubItem("Survival Curve", tabName = "survcurve"),
               menuSubItem("Cox Regression", tabName = "cox"),
               menuSubItem("Log-rank Test", tabName = "logrank")),
      menuItem("Regression Models", tabName = "regression", icon = icon("chart-line"),
               menuSubItem("Linear Regression", tabName = "linear"),
               menuSubItem("Logistic Regression", tabName = "logistic")),
      menuItem("Statistical Tests", tabName = "tests", icon = icon("flask"),
               menuSubItem("T-tests/Mann-Whitney", tabName = "ttest"),
               menuSubItem("ANOVA/Kruskal-Wallis", tabName = "anova"),
               menuSubItem("Chi-square Test", tabName = "chisq")),
      tags$div(
        style = "position: absolute; bottom: 0; width: 100%; padding: 10px; background-color: #222d32; color: white; text-align: center;",
        "Dev: Mudasir Mohammed Ibrahim",
        br(),
        "📧: mudassiribrahim30@gmail.com"
      )
    )
  ),
  dashboardBody(
    tabItems(
      # Data Upload
      tabItem(tabName = "data",
              fluidRow(
                box(title = "Upload Data", status = "primary", solidHeader = TRUE,
                    fileInput("file1", "Choose Data File",
                              accept = c(".csv", ".xlsx", ".xls", ".sav", ".dta")),
                    conditionalPanel(
                      condition = "input.file1.type == 'text/csv' || input.file1.type == 'text/comma-separated-values,text/plain'",
                      checkboxInput("header", "Header", TRUE),
                      radioButtons("sep", "Separator", choices = c(Comma = ",", Semicolon = ";", Tab = "\t"), selected = ","),
                      radioButtons("quote", "Quote", choices = c(None = "", "Double Quote" = '"', "Single Quote" = "'"), selected = '"')
                    ),
                    conditionalPanel(
                      condition = "input.file1.type == 'application/vnd.ms-excel' || input.file1.type == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'",
                      numericInput("sheet", "Sheet Number", value = 1)
                    )
                ),
                box(title = "Data Preview", status = "info", solidHeader = TRUE,
                    DTOutput("contents")
                )
              )
      ),
      
      # Frequencies
      tabItem(tabName = "freq",
              fluidRow(
                box(title = "Frequency Analysis", status = "primary", solidHeader = TRUE,
                    selectInput("freq_var", "Select Variable:", choices = NULL),
                    radioButtons("freq_type", "Output Type:", choices = c("Table", "Plot"), selected = "Table"),
                    conditionalPanel(
                      condition = "input.freq_type == 'Plot'",
                      colourpicker::colourInput("freq_color", "Select Color:", value = "steelblue"),
                      textInput("freq_title", "Plot Title:", value = "Frequency Plot"),
                      numericInput("freq_title_size", "Title Size:", value = 14),
                      numericInput("freq_x_size", "X-axis Label Size:", value = 12),
                      numericInput("freq_y_size", "Y-axis Label Size:", value = 12)
                    )
                ),
                box(title = "Results", status = "info", solidHeader = TRUE,
                    conditionalPanel(
                      condition = "input.freq_type == 'Table'",
                      DTOutput("freq_table")
                    ),
                    conditionalPanel(
                      condition = "input.freq_type == 'Plot'",
                      plotlyOutput("freq_plot")
                    ),
                    downloadButton("download_freq", "Download Results")
                )
              )
      ),
      
      # Central Tendency
      tabItem(tabName = "central",
              fluidRow(
                box(title = "Central Tendency", status = "primary", solidHeader = TRUE,
                    selectInput("central_var", "Select Numeric Variable:", choices = NULL),
                    radioButtons("central_type", "Output Type:", 
                                 choices = c("Table", "Boxplot", "Histogram"), 
                                 selected = "Table"),
                    conditionalPanel(
                      condition = "input.central_type != 'Table'",
                      colourpicker::colourInput("central_color", "Select Color:", value = "steelblue"),
                      textInput("central_title", "Plot Title:", value = "Distribution Plot"),
                      numericInput("central_title_size", "Title Size:", value = 14),
                      numericInput("central_x_size", "X-axis Label Size:", value = 12),
                      numericInput("central_y_size", "Y-axis Label Size:", value = 12),
                      conditionalPanel(
                        condition = "input.central_type == 'Histogram'",
                        numericInput("central_bins", "Number of Bins:", value = 30)
                      )
                    )
                ),
                box(title = "Results", status = "info", solidHeader = TRUE,
                    conditionalPanel(
                      condition = "input.central_type == 'Table'",
                      DTOutput("central_table")
                    ),
                    conditionalPanel(
                      condition = "input.central_type == 'Boxplot'",
                      plotlyOutput("central_boxplot")
                    ),
                    conditionalPanel(
                      condition = "input.central_type == 'Histogram'",
                      plotlyOutput("central_histogram")
                    ),
                    downloadButton("download_central", "Download Results")
                )
              )
      ),
      
      # Epidemic Curve
      tabItem(tabName = "epicurve",
              fluidRow(
                box(title = "Epidemic Curve", status = "primary", solidHeader = TRUE,
                    selectInput("epi_date", "Select Date Variable:", choices = NULL),
                    selectInput("epi_case", "Select Case Variable:", choices = NULL),
                    selectInput("epi_group", "Group By (optional):", choices = NULL),
                    radioButtons("epi_interval", "Time Interval:", 
                                 choices = c("Day", "Week", "Month", "Year"), selected = "Week"),
                    colourpicker::colourInput("epi_color", "Select Color:", value = "steelblue"),
                    textInput("epi_title", "Plot Title:", value = "Epidemic Curve"),
                    numericInput("epi_title_size", "Title Size:", value = 14),
                    numericInput("epi_x_size", "X-axis Label Size:", value = 12),
                    numericInput("epi_y_size", "Y-axis Label Size:", value = 12)
                ),
                box(title = "Epidemic Curve", status = "info", solidHeader = TRUE,
                    plotlyOutput("epi_curve"),
                    downloadButton("download_epicurve", "Download Plot")
                )
              )
      ),
      
      # Spatial Maps
      tabItem(tabName = "maps",
              fluidRow(
                box(title = "Spatial Distribution", status = "primary", solidHeader = TRUE,
                    selectInput("map_var", "Select Variable to Map:", choices = NULL),
                    selectInput("map_region", "Select Region Variable:", choices = NULL),
                    radioButtons("map_type", "Map Type:", choices = c("Choropleth", "Point"), selected = "Choropleth"),
                    colourpicker::colourInput("map_color", "Select Color:", value = "steelblue"),
                    textInput("map_title", "Plot Title:", value = "Spatial Distribution"),
                    numericInput("map_title_size", "Title Size:", value = 14),
                    numericInput("map_x_size", "X-axis Label Size:", value = 12),
                    numericInput("map_y_size", "Y-axis Label Size:", value = 12)
                ),
                box(title = "Map", status = "info", solidHeader = TRUE,
                    plotlyOutput("map_plot"),
                    downloadButton("download_map", "Download Map")
                )
              )
      ),
      
      # Demographics
      tabItem(tabName = "demo",
              fluidRow(
                box(title = "Demographic Breakdown", status = "primary", solidHeader = TRUE,
                    selectInput("demo_var", "Select Outcome Variable:", choices = NULL),
                    selectInput("demo_group", "Select Grouping Variable(s):", choices = NULL, multiple = TRUE)
                ),
                box(title = "Results", status = "info", solidHeader = TRUE,
                    div(style = 'overflow-x: scroll', DTOutput("demo_table")),
                    downloadButton("download_demo", "Download Table")
                )
              )
      ),
      
      # Incidence/Prevalence
      tabItem(tabName = "incprev",
              fluidRow(
                box(title = "Incidence/Prevalence", status = "primary", solidHeader = TRUE,
                    selectInput("ip_case", "Select Case Variable:", choices = NULL),
                    selectInput("ip_pop", "Select Population Variable:", choices = NULL),
                    selectInput("ip_time", "Select Time Variable:", choices = NULL),
                    radioButtons("ip_measure", "Measure:", 
                                 choices = c("Incidence Rate", "Cumulative Incidence", "Prevalence"), 
                                 selected = "Incidence Rate")
                ),
                box(title = "Results", status = "info", solidHeader = TRUE,
                    DTOutput("ip_table"),
                    downloadButton("download_ip", "Download Table")
                )
              )
      ),
      
      # Attack Rates
      tabItem(tabName = "attack",
              fluidRow(
                box(title = "Attack Rate", status = "primary", solidHeader = TRUE,
                    selectInput("ar_case", "Select Case Variable:", choices = NULL),
                    selectInput("ar_pop", "Select Population Variable:", choices = NULL),
                    selectInput("ar_group", "Group By (optional):", choices = NULL),
                    radioButtons("ar_format", "Output Format:", 
                                 choices = c("Table", "Detailed Table"), 
                                 selected = "Table")
                ),
                box(title = "Results", status = "info", solidHeader = TRUE,
                    conditionalPanel(
                      condition = "input.ar_format == 'Table'",
                      DTOutput("ar_table")
                    ),
                    conditionalPanel(
                      condition = "input.ar_format == 'Detailed Table'",
                      DTOutput("ar_detailed_table")
                    ),
                    downloadButton("download_ar", "Download Table")
                )
              )
      ),
      
      # Mortality Rates
      tabItem(tabName = "mortality",
              fluidRow(
                box(title = "Mortality Rates", status = "primary", solidHeader = TRUE,
                    selectInput("mort_case", "Select Death Variable:", choices = NULL),
                    selectInput("mort_pop", "Select Population Variable:", choices = NULL),
                    radioButtons("mort_measure", "Measure:", 
                                 choices = c("Mortality Rate", "Case Fatality Rate"), 
                                 selected = "Mortality Rate")
                ),
                box(title = "Results", status = "info", solidHeader = TRUE,
                    DTOutput("mort_table"),
                    downloadButton("download_mort", "Download Table")
                )
              )
      ),
      
      # Life Table Analysis
      tabItem(tabName = "lifetable",
              fluidRow(
                box(title = "Life Table Analysis", status = "primary", solidHeader = TRUE,
                    selectInput("lt_time", "Select Time Variable:", choices = NULL),
                    selectInput("lt_event", "Select Event Variable:", choices = NULL),
                    textInput("lt_breaks", "Time Breaks (comma separated):", value = "0,10,20,30,40,50,60,70,80"),
                    textInput("lt_title", "Table Title:", value = "Life Table Analysis")
                ),
                box(title = "Life Table Results", status = "info", solidHeader = TRUE,
                    div(style = 'overflow-x: scroll', DTOutput("lifetable_results")),
                    downloadButton("download_lifetable", "Download Table")
                )
              )
      ),
      
      # Risk Ratios
      tabItem(tabName = "risk",
              fluidRow(
                box(title = "Risk Ratio", status = "primary", solidHeader = TRUE,
                    selectInput("rr_outcome", "Select Outcome Variable:", choices = NULL),
                    selectInput("rr_exposure", "Select Exposure Variable:", choices = NULL),
                    numericInput("rr_conf", "Confidence Level:", value = 0.95, min = 0.5, max = 0.99, step = 0.01),
                    textInput("rr_exposed_label", "Exposed Group Label:", value = "Exposed"),
                    textInput("rr_unexposed_label", "Unexposed Group Label:", value = "Unexposed")
                ),
                box(title = "Results", status = "info", solidHeader = TRUE,
                    verbatimTextOutput("rr_results"),
                    DTOutput("rr_table"),
                    downloadButton("download_rr", "Download Table")
                )
              )
      ),
      
      # Odds Ratios
      tabItem(tabName = "odds",
              fluidRow(
                box(title = "Odds Ratio", status = "primary", solidHeader = TRUE,
                    selectInput("or_outcome", "Select Outcome Variable:", choices = NULL),
                    selectInput("or_exposure", "Select Exposure Variable:", choices = NULL),
                    numericInput("or_conf", "Confidence Level:", value = 0.95, min = 0.5, max = 0.99, step = 0.01),
                    textInput("or_exposed_label", "Exposed Group Label:", value = "Exposed"),
                    textInput("or_unexposed_label", "Unexposed Group Label:", value = "Unexposed")
                ),
                box(title = "Results", status = "info", solidHeader = TRUE,
                    verbatimTextOutput("or_results"),
                    DTOutput("or_table"),
                    downloadButton("download_or", "Download Table")
                )
              )
      ),
      
      # Attributable Risk
      tabItem(tabName = "ar",
              fluidRow(
                box(title = "Attributable Risk", status = "primary", solidHeader = TRUE,
                    selectInput("ar_outcome", "Select Outcome Variable:", choices = NULL),
                    selectInput("ar_exposure", "Select Exposure Variable:", choices = NULL),
                    numericInput("ar_conf", "Confidence Level:", value = 0.95, min = 0.5, max = 0.99, step = 0.01),
                    checkboxInput("ar_paf", "Calculate Population Attributable Fraction", value = FALSE),
                    textInput("ar_exposed_label", "Exposed Group Label:", value = "Exposed"),
                    textInput("ar_unexposed_label", "Unexposed Group Label:", value = "Unexposed")
                ),
                box(title = "Results", status = "info", solidHeader = TRUE,
                    verbatimTextOutput("ar_results"),
                    DTOutput("ar_table"),
                    downloadButton("download_arisk", "Download Table")
                )
              )
      ),
      
      # Poisson Regression
      tabItem(tabName = "poisson",
              fluidRow(
                box(title = "Poisson Regression", status = "primary", solidHeader = TRUE,
                    selectInput("poisson_outcome", "Select Outcome Count Variable:", choices = NULL),
                    selectInput("poisson_offset", "Select Offset Variable (log transform will be applied):", choices = NULL),
                    selectInput("poisson_vars", "Select Predictor Variables:", choices = NULL, multiple = TRUE),
                    uiOutput("poisson_ref_ui"),
                    numericInput("poisson_conf", "Confidence Level:", value = 0.95, min = 0.5, max = 0.99, step = 0.01),
                    checkboxInput("poisson_irr", "Show Incidence Rate Ratios (IRR)", value = TRUE)
                ),
                box(title = "Poisson Regression Results", status = "info", solidHeader = TRUE,
                    verbatimTextOutput("poisson_summary"),
                    DTOutput("poisson_table"),
                    downloadButton("download_poisson", "Download Results")
                )
              )
      ),
      
      # Survival Curve (Kaplan-Meier)
      tabItem(tabName = "survcurve",
              fluidRow(
                box(title = "Survival Analysis", status = "primary", solidHeader = TRUE,
                    selectInput("surv_time", "Select Time Variable:", choices = NULL),
                    selectInput("surv_event", "Select Event Variable:", choices = NULL),
                    selectInput("surv_group", "Group By (optional):", choices = NULL),
                    numericInput("surv_conf", "Confidence Level:", value = 0.95, min = 0.5, max = 0.99, step = 0.01),
                    checkboxInput("surv_median", "Show Median Survival", value = TRUE),
                    checkboxInput("surv_mean", "Show Mean Survival", value = FALSE),
                    checkboxInput("surv_hazard", "Show Hazard Ratio", value = FALSE),
                    checkboxInput("surv_ph", "Test Proportional Hazards", value = FALSE),
                    colourpicker::colourInput("surv_color1", "Color for Group 1:", value = "#E41A1C"),
                    colourpicker::colourInput("surv_color2", "Color for Group 2:", value = "#377EB8"),
                    textInput("surv_title", "Plot Title:", value = "Time-to-Event Survival Curve"),
                    numericInput("surv_title_size", "Title Size:", value = 14),
                    numericInput("surv_x_size", "X-axis Label Size:", value = 12),
                    numericInput("surv_y_size", "Y-axis Label Size:", value = 12),
                    numericInput("surv_break_time", "Break Time By:", value = NULL, min = 1),
                    textInput("surv_xlab", "X-axis Label:", value = "Time"),
                    textInput("surv_ylab", "Y-axis Label:", value = "Survival Probability"),
                    textInput("surv_legend_title", "Legend Title:", value = "Group"),
                    textInput("surv_legend_labels", "Legend Labels (comma separated):", value = "")
                ),
                box(title = "Survival Analysis Results", status = "info", solidHeader = TRUE,
                    verbatimTextOutput("surv_summary"),
                    plotOutput("surv_plot"),
                    downloadButton("download_surv", "Download Plot")
                )
              )
      ),
      
      # Log-Rank Test
      tabItem(tabName = "logrank",
              fluidRow(
                box(title = "Log-Rank Test", status = "primary", solidHeader = TRUE,
                    selectInput("lr_time", "Select Time Variable:", choices = NULL),
                    selectInput("lr_event", "Select Event Variable:", choices = NULL),
                    selectInput("lr_group", "Select Grouping Variable:", choices = NULL)
                ),
                box(title = "Results", status = "info", solidHeader = TRUE,
                    verbatimTextOutput("logrank_results"),
                    downloadButton("download_logrank", "Download Results")
                )
              )
      ),
      
      # Cox Regression
      tabItem(tabName = "cox",
              fluidRow(
                box(title = "Cox Regression", status = "primary", solidHeader = TRUE,
                    selectInput("cox_time", "Select Time Variable:", choices = NULL),
                    selectInput("cox_event", "Select Event Variable:", choices = NULL),
                    selectInput("cox_vars", "Select Predictor Variables:", choices = NULL, multiple = TRUE),
                    uiOutput("cox_ref_ui"),
                    numericInput("cox_conf", "Confidence Level:", value = 0.95, min = 0.5, max = 0.99, step = 0.01)
                ),
                box(title = "Results", status = "info", solidHeader = TRUE,
                    verbatimTextOutput("cox_summary"),
                    DTOutput("cox_table"),
                    downloadButton("download_cox", "Download Results")
                )
              )
      ),
      
      # Linear Regression
      tabItem(tabName = "linear",
              fluidRow(
                box(
                  title = "Linear Regression", status = "primary", solidHeader = TRUE, width = 6,
                  selectInput("linear_outcome", "Select Outcome Variable:", choices = NULL),
                  selectInput("linear_vars", "Select Predictor Variables:", choices = NULL, multiple = TRUE),
                  uiOutput("linear_ref_ui"),
                  checkboxInput("linear_std", "Show Standardized Estimates", value = TRUE)
                ),
                box(
                  title = "Results", status = "info", solidHeader = TRUE, width = 6,
                  div(style = "overflow-x: auto;",
                      verbatimTextOutput("linear_summary")
                  ),
                  div(style = "overflow-x: auto;",
                      DTOutput("linear_table")
                  ),
                  downloadButton("download_linear", "Download Results")
                )
              )
      ),
      
      
      # Logistic Regression
      tabItem(tabName = "logistic",
              fluidRow(
                box(title = "Logistic Regression", status = "primary", solidHeader = TRUE,
                    radioButtons("logistic_type", "Regression Type:",
                                 choices = c("Binary" = "binary", 
                                             "Ordinal" = "ordinal",
                                             "Multinomial" = "multinomial"),
                                 selected = "binary"),
                    selectInput("logistic_outcome", "Select Outcome Variable:", choices = NULL),
                    selectInput("logistic_vars", "Select Predictor Variables:", choices = NULL, multiple = TRUE),
                    uiOutput("logistic_ref_ui"),
                    checkboxInput("logistic_or", "Show Odds Ratios", value = TRUE)
                ),
                box(title = "Results", status = "info", solidHeader = TRUE,
                    verbatimTextOutput("logistic_summary"),
                    DTOutput("logistic_table"),
                    downloadButton("download_logistic", "Download Results")
                )
              )
      ),
      
      # T-tests/Mann-Whitney
      tabItem(tabName = "ttest",
              fluidRow(
                box(title = "T-test/Mann-Whitney", status = "primary", solidHeader = TRUE,
                    radioButtons("ttest_type", "Test Type:",
                                 choices = c("One Sample t-test" = "one.sample",
                                             "Two Sample t-test" = "two.sample",
                                             "Paired t-test" = "paired",
                                             "Mann-Whitney U Test" = "mannwhitney"),
                                 selected = "two.sample"),
                    selectInput("ttest_var", "Select Variable:", choices = NULL),
                    conditionalPanel(
                      condition = "input.ttest_type != 'one.sample'",
                      selectInput("ttest_group", "Select Grouping Variable:", choices = NULL)
                    ),
                    conditionalPanel(
                      condition = "input.ttest_type == 'one.sample'",
                      numericInput("ttest_mu", "Comparison Mean:", value = 0)
                    ),
                    conditionalPanel(
                      condition = "input.ttest_type == 'two.sample'",
                      checkboxInput("ttest_var_equal", "Assume Equal Variances", value = FALSE)
                    )
                ),
                box(title = "Results", status = "info", solidHeader = TRUE,
                    verbatimTextOutput("ttest_results"),
                    downloadButton("download_ttest", "Download Results")
                )
              )
      ),
      
      # ANOVA/Kruskal-Wallis
      tabItem(tabName = "anova",
              fluidRow(
                box(title = "ANOVA/Kruskal-Wallis", status = "primary", solidHeader = TRUE,
                    radioButtons("anova_type", "Test Type:",
                                 choices = c("One-way ANOVA" = "anova",
                                             "Kruskal-Wallis" = "kruskal"),
                                 selected = "anova"),
                    selectInput("anova_var", "Select Variable:", choices = NULL),
                    selectInput("anova_group", "Select Grouping Variable:", choices = NULL),
                    checkboxInput("anova_posthoc", "Perform Post-hoc Tests", value = TRUE)
                ),
                box(title = "Results", status = "info", solidHeader = TRUE,
                    verbatimTextOutput("anova_results"),
                    conditionalPanel(
                      condition = "input.anova_posthoc",
                      verbatimTextOutput("posthoc_results")
                    ),
                    downloadButton("download_anova", "Download Results")
                )
              )
      ),
      
      # Chi-square Test
      tabItem(tabName = "chisq",
              fluidRow(
                box(title = "Chi-square Test", status = "primary", solidHeader = TRUE,
                    selectInput("chisq_var1", "Select First Variable:", choices = NULL),
                    selectInput("chisq_var2", "Select Second Variable:", choices = NULL),
                    checkboxInput("chisq_expected", "Show Expected Counts", value = FALSE),
                    checkboxInput("chisq_residuals", "Show Residuals", value = FALSE)
                ),
                box(title = "Results", status = "info", solidHeader = TRUE,
                    verbatimTextOutput("chisq_results"),
                    DTOutput("chisq_table"),
                    downloadButton("download_chisq", "Download Results")
                )
              )
      )
    )
  )
)

# Server function
server <- function(input, output, session) {
  
  # Reactive data
  data <- reactive({
    req(input$file1)
    
    ext <- tools::file_ext(input$file1$name)
    
    df <- switch(ext,
                 csv = read.csv(input$file1$datapath,
                                header = input$header,
                                sep = input$sep,
                                quote = input$quote),
                 xls = read_excel(input$file1$datapath, sheet = input$sheet),
                 xlsx = read_excel(input$file1$datapath, sheet = input$sheet),
                 sav = read_sav(input$file1$datapath),
                 dta = read_dta(input$file1$datapath),
                 validate("Invalid file; Please upload a .csv, .xls, .xlsx, .sav, or .dta file")
    )
    
    # Convert to data.frame if it's a tibble
    if (inherits(df, "tbl_df")) {
      df <- as.data.frame(df)
    }
    
    # Update all select inputs
    updateSelectInput(session, "freq_var", choices = names(df))
    updateSelectInput(session, "central_var", choices = names(df)[sapply(df, is.numeric)])
    updateSelectInput(session, "epi_date", choices = names(df)[sapply(df, function(x) inherits(x, "Date") | is.character(x))])
    updateSelectInput(session, "epi_case", choices = names(df))
    updateSelectInput(session, "epi_group", choices = c("None", names(df)))
    updateSelectInput(session, "map_var", choices = names(df))
    updateSelectInput(session, "map_region", choices = names(df))
    updateSelectInput(session, "demo_var", choices = names(df))
    updateSelectInput(session, "demo_group", choices = names(df))
    updateSelectInput(session, "ip_case", choices = names(df))
    updateSelectInput(session, "ip_pop", choices = names(df)[sapply(df, is.numeric)])
    updateSelectInput(session, "ip_time", choices = names(df))
    updateSelectInput(session, "ar_case", choices = names(df))
    updateSelectInput(session, "ar_pop", choices = names(df)[sapply(df, is.numeric)])
    updateSelectInput(session, "ar_group", choices = c("None", names(df)))
    updateSelectInput(session, "mort_case", choices = names(df))
    updateSelectInput(session, "mort_pop", choices = names(df)[sapply(df, is.numeric)])
    updateSelectInput(session, "lt_time", choices = names(df)[sapply(df, is.numeric)])
    updateSelectInput(session, "lt_event", choices = names(df))
    updateSelectInput(session, "rr_outcome", choices = names(df))
    updateSelectInput(session, "rr_exposure", choices = names(df))
    updateSelectInput(session, "or_outcome", choices = names(df))
    updateSelectInput(session, "or_exposure", choices = names(df))
    updateSelectInput(session, "ar_outcome", choices = names(df))
    updateSelectInput(session, "ar_exposure", choices = names(df))
    updateSelectInput(session, "poisson_outcome", choices = names(df)[sapply(df, is.numeric)])
    updateSelectInput(session, "poisson_offset", choices = names(df)[sapply(df, is.numeric)])
    updateSelectInput(session, "poisson_vars", choices = names(df))
    updateSelectInput(session, "surv_time", choices = names(df)[sapply(df, is.numeric)])
    updateSelectInput(session, "surv_event", choices = names(df))
    updateSelectInput(session, "surv_group", choices = c("None", names(df)))
    updateSelectInput(session, "lr_time", choices = names(df)[sapply(df, is.numeric)])
    updateSelectInput(session, "lr_event", choices = names(df))
    updateSelectInput(session, "lr_group", choices = names(df))
    updateSelectInput(session, "cox_time", choices = names(df)[sapply(df, is.numeric)])
    updateSelectInput(session, "cox_event", choices = names(df))
    updateSelectInput(session, "cox_vars", choices = names(df))
    updateSelectInput(session, "linear_outcome", choices = names(df)[sapply(df, is.numeric)])
    updateSelectInput(session, "linear_vars", choices = names(df))
    updateSelectInput(session, "logistic_outcome", choices = names(df))
    updateSelectInput(session, "logistic_vars", choices = names(df))
    updateSelectInput(session, "ttest_var", choices = names(df)[sapply(df, is.numeric)])
    updateSelectInput(session, "ttest_group", choices = names(df))
    updateSelectInput(session, "anova_var", choices = names(df)[sapply(df, is.numeric)])
    updateSelectInput(session, "anova_group", choices = names(df))
    updateSelectInput(session, "chisq_var1", choices = names(df))
    updateSelectInput(session, "chisq_var2", choices = names(df))
    
    df
  })
  
  # Data preview
  output$contents <- renderDT({
    datatable(data(), options = list(scrollX = TRUE))
  })
  
  # Frequency analysis
  output$freq_table <- renderDT({
    req(input$freq_var)
    df <- data()
    freq_table <- as.data.frame(table(df[[input$freq_var]]))
    names(freq_table) <- c(input$freq_var, "Frequency")
    freq_table$Proportion <- freq_table$Frequency / sum(freq_table$Frequency)
    datatable(freq_table)
  })
  
  output$freq_plot <- renderPlotly({
    req(input$freq_var)
    df <- data()
    
    p <- ggplot(df, aes(x = .data[[input$freq_var]])) +
      geom_bar(fill = input$freq_color) +
      labs(title = input$freq_title, x = input$freq_var, y = "Count") +
      theme_minimal() +
      theme(
        plot.title = element_text(size = input$freq_title_size),
        axis.title.x = element_text(size = input$freq_x_size),
        axis.title.y = element_text(size = input$freq_y_size)
      )
    
    ggplotly(p)
  })
  
  output$download_freq <- downloadHandler(
    filename = function() {
      if(input$freq_type == "Table") {
        "frequency_table.docx"
      } else {
        "frequency_plot.png"
      }
    },
    content = function(file) {
      if(input$freq_type == "Table") {
        df <- data()
        freq_table <- as.data.frame(table(df[[input$freq_var]]))
        names(freq_table) <- c(input$freq_var, "Frequency")
        freq_table$Proportion <- freq_table$Frequency / sum(freq_table$Frequency)
        
        ft <- flextable(freq_table)
        doc <- read_docx() %>% 
          body_add_flextable(ft)
        print(doc, target = file)
      } else {
        df <- data()
        p <- ggplot(df, aes(x = .data[[input$freq_var]])) +
          geom_bar(fill = input$freq_color) +
          labs(title = input$freq_title, x = input$freq_var, y = "Count") +
          theme_minimal() +
          theme(
            plot.title = element_text(size = input$freq_title_size),
            axis.title.x = element_text(size = input$freq_x_size),
            axis.title.y = element_text(size = input$freq_y_size)
          )
        ggsave(file, plot = p, device = "png", width = 8, height = 6)
      }
    }
  )
  
  # Central tendency
  output$central_table <- renderDT({
    req(input$central_var)
    df <- data()
    var <- df[[input$central_var]]
    
    # Calculate mode
    get_mode <- function(x) {
      ux <- unique(x)
      ux[which.max(tabulate(match(x, ux)))]
    }
    
    central_table <- data.frame(
      Measure = c("Mean", "Median", "SD", "Min", "Max", "IQR", "Mode"),
      Value = c(mean(var, na.rm = TRUE), 
                median(var, na.rm = TRUE),
                sd(var, na.rm = TRUE),
                min(var, na.rm = TRUE),
                max(var, na.rm = TRUE),
                IQR(var, na.rm = TRUE),
                get_mode(var))
    )
    datatable(central_table)
  })
  
  output$central_boxplot <- renderPlotly({
    req(input$central_var)
    df <- data()
    
    p <- ggplot(df, aes(y = .data[[input$central_var]])) +
      geom_boxplot(fill = input$central_color) +
      labs(title = input$central_title, y = input$central_var) +
      theme_minimal() +
      theme(
        plot.title = element_text(size = input$central_title_size),
        axis.title.x = element_text(size = input$central_x_size),
        axis.title.y = element_text(size = input$central_y_size)
      )
    
    ggplotly(p)
  })
  
  output$central_histogram <- renderPlotly({
    req(input$central_var)
    df <- data()
    
    p <- ggplot(df, aes(x = .data[[input$central_var]])) +
      geom_histogram(fill = input$central_color, bins = input$central_bins) +
      labs(title = input$central_title, x = input$central_var, y = "Count") +
      theme_minimal() +
      theme(
        plot.title = element_text(size = input$central_title_size),
        axis.title.x = element_text(size = input$central_x_size),
        axis.title.y = element_text(size = input$central_y_size)
      )
    
    ggplotly(p)
  })
  
  output$download_central <- downloadHandler(
    filename = function() {
      if(input$central_type == "Table") {
        "central_tendency_table.docx"
      } else if(input$central_type == "Boxplot") {
        "central_tendency_boxplot.png"
      } else {
        "central_tendency_histogram.png"
      }
    },
    content = function(file) {
      df <- data()
      var <- df[[input$central_var]]
      
      if(input$central_type == "Table") {
        get_mode <- function(x) {
          ux <- unique(x)
          ux[which.max(tabulate(match(x, ux)))]
        }
        
        central_table <- data.frame(
          Measure = c("Mean", "Median", "SD", "Min", "Max", "IQR", "Mode"),
          Value = c(mean(var, na.rm = TRUE), 
                    median(var, na.rm = TRUE),
                    sd(var, na.rm = TRUE),
                    min(var, na.rm = TRUE),
                    max(var, na.rm = TRUE),
                    IQR(var, na.rm = TRUE),
                    get_mode(var))
        )
        
        ft <- flextable(central_table)
        doc <- read_docx() %>% 
          body_add_flextable(ft)
        print(doc, target = file)
      } else if(input$central_type == "Boxplot") {
        p <- ggplot(df, aes(y = .data[[input$central_var]])) +
          geom_boxplot(fill = input$central_color) +
          labs(title = input$central_title, y = input$central_var) +
          theme_minimal() +
          theme(
            plot.title = element_text(size = input$central_title_size),
            axis.title.x = element_text(size = input$central_x_size),
            axis.title.y = element_text(size = input$central_y_size)
          )
        ggsave(file, plot = p, device = "png", width = 8, height = 6)
      } else {
        p <- ggplot(df, aes(x = .data[[input$central_var]])) +
          geom_histogram(fill = input$central_color, bins = input$central_bins) +
          labs(title = input$central_title, x = input$central_var, y = "Count") +
          theme_minimal() +
          theme(
            plot.title = element_text(size = input$central_title_size),
            axis.title.x = element_text(size = input$central_x_size),
            axis.title.y = element_text(size = input$central_y_size)
          )
        ggsave(file, plot = p, device = "png", width = 8, height = 6)
      }
    }
  )
  
  # Epidemic curve
  output$epi_curve <- renderPlotly({
    req(input$epi_date, input$epi_case)
    df <- data()
    
    # Convert date variable to Date class if it's character
    if(is.character(df[[input$epi_date]])) {
      df[[input$epi_date]] <- as.Date(df[[input$epi_date]])
    }
    
    # Aggregate data by time interval
    if(input$epi_interval == "Day") {
      df$time_group <- as.Date(df[[input$epi_date]])
    } else if(input$epi_interval == "Week") {
      df$time_group <- cut(df[[input$epi_date]], breaks = "week")
    } else if(input$epi_interval == "Month") {
      df$time_group <- cut(df[[input$epi_date]], breaks = "month")
    } else {
      df$time_group <- cut(df[[input$epi_date]], breaks = "year")
    }
    
    # Group data
    if(input$epi_group != "None") {
      epi_data <- df %>%
        group_by(time_group, .data[[input$epi_group]]) %>%
        summarise(cases = sum(.data[[input$epi_case]]), .groups = "drop")
      
      p <- ggplot(epi_data, aes(x = time_group, y = cases, fill = .data[[input$epi_group]])) +
        geom_col() +
        scale_fill_brewer(palette = "Set1")
    } else {
      epi_data <- df %>%
        group_by(time_group) %>%
        summarise(cases = sum(.data[[input$epi_case]]), .groups = "drop")
      
      p <- ggplot(epi_data, aes(x = time_group, y = cases)) +
        geom_col(fill = input$epi_color)
    }
    
    p <- p +
      labs(title = input$epi_title, x = "Date", y = "Number of Cases") +
      theme_minimal() +
      theme(
        plot.title = element_text(size = input$epi_title_size),
        axis.title.x = element_text(size = input$epi_x_size),
        axis.title.y = element_text(size = input$epi_y_size),
        axis.text.x = element_text(angle = 45, hjust = 1)
      )
    
    ggplotly(p)
  })
  
  output$download_epicurve <- downloadHandler(
    filename = function() {
      "epidemic_curve.png"
    },
    content = function(file) {
      df <- data()
      
      # Convert date variable to Date class if it's character
      if(is.character(df[[input$epi_date]])) {
        df[[input$epi_date]] <- as.Date(df[[input$epi_date]])
      }
      
      # Aggregate data by time interval
      if(input$epi_interval == "Day") {
        df$time_group <- as.Date(df[[input$epi_date]])
      } else if(input$epi_interval == "Week") {
        df$time_group <- cut(df[[input$epi_date]], breaks = "week")
      } else if(input$epi_interval == "Month") {
        df$time_group <- cut(df[[input$epi_date]], breaks = "month")
      } else {
        df$time_group <- cut(df[[input$epi_date]], breaks = "year")
      }
      
      # Group data
      if(input$epi_group != "None") {
        epi_data <- df %>%
          group_by(time_group, .data[[input$epi_group]]) %>%
          summarise(cases = sum(.data[[input$epi_case]]), .groups = "drop")
        
        p <- ggplot(epi_data, aes(x = time_group, y = cases, fill = .data[[input$epi_group]])) +
          geom_col() +
          scale_fill_brewer(palette = "Set1")
      } else {
        epi_data <- df %>%
          group_by(time_group) %>%
          summarise(cases = sum(.data[[input$epi_case]]), .groups = "drop")
        
        p <- ggplot(epi_data, aes(x = time_group, y = cases)) +
          geom_col(fill = input$epi_color)
      }
      
      p <- p +
        labs(title = input$epi_title, x = "Date", y = "Number of Cases") +
        theme_minimal() +
        theme(
          plot.title = element_text(size = input$epi_title_size),
          axis.title.x = element_text(size = input$epi_x_size),
          axis.title.y = element_text(size = input$epi_y_size),
          axis.text.x = element_text(angle = 45, hjust = 1)
        )
      
      ggsave(file, plot = p, device = "png", width = 10, height = 6)
    }
  )
  # Spatial maps (simplified - would need spatial data for full implementation)
  output$map_plot <- renderPlotly({
    req(input$map_var, input$map_region)
    df <- data()
    
    if(input$map_type == "Choropleth") {
      # Simplified choropleth - would need spatial data for real implementation
      region_counts <- df %>%
        group_by(.data[[input$map_region]]) %>%
        summarise(value = mean(.data[[input$map_var]], na.rm = TRUE))
      
      p <- ggplot(region_counts, aes(x = .data[[input$map_region]], y = value, fill = value)) +
        geom_col() +
        scale_fill_gradient(low = "blue", high = "red") +
        labs(title = input$map_title, 
             x = input$map_region, y = input$map_var) +
        theme_minimal() +
        theme(
          plot.title = element_text(size = input$map_title_size),
          axis.title.x = element_text(size = input$map_x_size),
          axis.title.y = element_text(size = input$map_y_size),
          axis.text.x = element_text(angle = 45, hjust = 1)
        )
    } else {
      # Point map - simplified
      p <- ggplot(df, aes(x = .data[[input$map_region]], y = .data[[input$map_var]])) +
        geom_point(color = input$map_color, size = 3) +
        labs(title = input$map_title, 
             x = input$map_region, y = input$map_var) +
        theme_minimal() +
        theme(
          plot.title = element_text(size = input$map_title_size),
          axis.title.x = element_text(size = input$map_x_size),
          axis.title.y = element_text(size = input$map_y_size)
        )
    }
    
    ggplotly(p)
  })
  
  output$download_map <- downloadHandler(
    filename = "spatial_map.png",
    content = function(file) {
      df <- data()
      
      if(input$map_type == "Choropleth") {
        region_counts <- df %>%
          group_by(.data[[input$map_region]]) %>%
          summarise(value = mean(.data[[input$map_var]], na.rm = TRUE))
        
        p <- ggplot(region_counts, aes(x = .data[[input$map_region]], y = value, fill = value)) +
          geom_col() +
          scale_fill_gradient(low = "blue", high = "red") +
          labs(title = input$map_title, 
               x = input$map_region, y = input$map_var) +
          theme_minimal() +
          theme(
            plot.title = element_text(size = input$map_title_size),
            axis.title.x = element_text(size = input$map_x_size),
            axis.title.y = element_text(size = input$map_y_size),
            axis.text.x = element_text(angle = 45, hjust = 1)
          )
      } else {
        p <- ggplot(df, aes(x = .data[[input$map_region]], y = .data[[input$map_var]])) +
          geom_point(color = input$map_color, size = 3) +
          labs(title = input$map_title, 
               x = input$map_region, y = input$map_var) +
          theme_minimal() +
          theme(
            plot.title = element_text(size = input$map_title_size),
            axis.title.x = element_text(size = input$map_x_size),
            axis.title.y = element_text(size = input$map_y_size)
          )
      }
      
      ggsave(file, plot = p, device = "png", width = 10, height = 6)
    }
  )
  
  
  # Demographic breakdown
  output$demo_table <- renderDT({
    req(input$demo_var, input$demo_group)
    df <- data()
    
    if(length(input$demo_group) == 1) {
      demo_table <- df %>%
        group_by(.data[[input$demo_group]]) %>%
        summarise(
          Count = n(),
          Mean = mean(.data[[input$demo_var]], na.rm = TRUE),
          SD = sd(.data[[input$demo_var]], na.rm = TRUE),
          Median = median(.data[[input$demo_var]], na.rm = TRUE),
          Min = min(.data[[input$demo_var]], na.rm = TRUE),
          Max = max(.data[[input$demo_var]], na.rm = TRUE),
          Mode = {
            ux <- unique(.data[[input$demo_var]])
            ux[which.max(tabulate(match(.data[[input$demo_var]], ux)))]
          }
        )
    } else {
      demo_table <- df %>%
        group_by(across(all_of(input$demo_group))) %>%
        summarise(
          Count = n(),
          Mean = mean(.data[[input$demo_var]], na.rm = TRUE),
          SD = sd(.data[[input$demo_var]], na.rm = TRUE),
          Median = median(.data[[input$demo_var]], na.rm = TRUE),
          Min = min(.data[[input$demo_var]], na.rm = TRUE),
          Max = max(.data[[input$demo_var]], na.rm = TRUE),
          Mode = {
            ux <- unique(.data[[input$demo_var]])
            ux[which.max(tabulate(match(.data[[input$demo_var]], ux)))]
          }
        )
    }
    
    datatable(demo_table)
  })
  
  output$download_demo <- downloadHandler(
    filename = "demographic_table.docx",
    content = function(file) {
      df <- data()
      
      if(length(input$demo_group) == 1) {
        demo_table <- df %>%
          group_by(.data[[input$demo_group]]) %>%
          summarise(
            Count = n(),
            Mean = mean(.data[[input$demo_var]], na.rm = TRUE),
            SD = sd(.data[[input$demo_var]], na.rm = TRUE),
            Median = median(.data[[input$demo_var]], na.rm = TRUE),
            Min = min(.data[[input$demo_var]], na.rm = TRUE),
            Max = max(.data[[input$demo_var]], na.rm = TRUE),
            Mode = {
              ux <- unique(.data[[input$demo_var]])
              ux[which.max(tabulate(match(.data[[input$demo_var]], ux)))]
            }
          )
      } else {
        demo_table <- df %>%
          group_by(across(all_of(input$demo_group))) %>%
          summarise(
            Count = n(),
            Mean = mean(.data[[input$demo_var]], na.rm = TRUE),
            SD = sd(.data[[input$demo_var]], na.rm = TRUE),
            Median = median(.data[[input$demo_var]], na.rm = TRUE),
            Min = min(.data[[input$demo_var]], na.rm = TRUE),
            Max = max(.data[[input$demo_var]], na.rm = TRUE),
            Mode = {
              ux <- unique(.data[[input$demo_var]])
              ux[which.max(tabulate(match(.data[[input$demo_var]], ux)))]
            }
          )
      }
      
      ft <- flextable(demo_table)
      doc <- read_docx() %>% 
        body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  # Incidence/Prevalence
  output$ip_table <- renderDT({
    req(input$ip_case, input$ip_pop)
    df <- data()
    
    if(input$ip_measure == "Incidence Rate") {
      req(input$ip_time)
      total_cases <- sum(df[[input$ip_case]], na.rm = TRUE)
      total_pop <- sum(df[[input$ip_pop]], na.rm = TRUE)
      total_time <- sum(df[[input$ip_time]], na.rm = TRUE)
      
      rate <- (total_cases / total_pop) * (1 / total_time) * 1000  # per 1000 person-time
      
      result <- data.frame(
        Measure = "Incidence Rate",
        Cases = total_cases,
        Population = total_pop,
        PersonTime = total_time,
        Rate = rate,
        Units = "per 1000 person-time"
      )
    } else if(input$ip_measure == "Cumulative Incidence") {
      total_cases <- sum(df[[input$ip_case]], na.rm = TRUE)
      total_pop <- sum(df[[input$ip_pop]], na.rm = TRUE)
      
      ci <- (total_cases / total_pop) * 1000  # per 1000 population
      
      result <- data.frame(
        Measure = "Cumulative Incidence",
        Cases = total_cases,
        Population = total_pop,
        Rate = ci,
        Units = "per 1000 population"
      )
    } else {
      # Prevalence
      total_cases <- sum(df[[input$ip_case]], na.rm = TRUE)
      total_pop <- sum(df[[input$ip_pop]], na.rm = TRUE)
      
      prevalence <- (total_cases / total_pop) * 100  # percentage
      
      result <- data.frame(
        Measure = "Prevalence",
        Cases = total_cases,
        Population = total_pop,
        Rate = prevalence,
        Units = "percent"
      )
    }
    
    datatable(result)
  })
  
  output$download_ip <- downloadHandler(
    filename = function() {
      "incidence_prevalence.docx"
    },
    content = function(file) {
      df <- data()
      
      if(input$ip_measure == "Incidence Rate") {
        req(input$ip_time)
        total_cases <- sum(df[[input$ip_case]], na.rm = TRUE)
        total_pop <- sum(df[[input$ip_pop]], na.rm = TRUE)
        total_time <- sum(df[[input$ip_time]], na.rm = TRUE)
        
        rate <- (total_cases / total_pop) * (1 / total_time) * 1000
        
        result <- data.frame(
          Measure = "Incidence Rate",
          Cases = total_cases,
          Population = total_pop,
          PersonTime = total_time,
          Rate = rate,
          Units = "per 1000 person-time"
        )
      } else if(input$ip_measure == "Cumulative Incidence") {
        total_cases <- sum(df[[input$ip_case]], na.rm = TRUE)
        total_pop <- sum(df[[input$ip_pop]], na.rm = TRUE)
        
        ci <- (total_cases / total_pop) * 1000
        
        result <- data.frame(
          Measure = "Cumulative Incidence",
          Cases = total_cases,
          Population = total_pop,
          Rate = ci,
          Units = "per 1000 population"
        )
      } else {
        total_cases <- sum(df[[input$ip_case]], na.rm = TRUE)
        total_pop <- sum(df[[input$ip_pop]], na.rm = TRUE)
        
        prevalence <- (total_cases / total_pop) * 100
        
        result <- data.frame(
          Measure = "Prevalence",
          Cases = total_cases,
          Population = total_pop,
          Rate = prevalence,
          Units = "percent"
        )
      }
      
      ft <- flextable(result)
      doc <- read_docx() %>% 
        body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  # Attack Rates Table
  output$ar_table <- renderDT({
    req(input$ar_case, input$ar_pop)
    df <- data()
    
    if (input$ar_group != "None") {
      ar_data <- df %>%
        group_by(.data[[input$ar_group]]) %>%
        summarise(
          Cases = sum(.data[[input$ar_case]], na.rm = TRUE),
          Population = sum(.data[[input$ar_pop]], na.rm = TRUE),
          AttackRate = (Cases / Population) * 100,
          .groups = "drop"
        )
    } else {
      ar_data <- data.frame(
        Cases = sum(df[[input$ar_case]], na.rm = TRUE),
        Population = sum(df[[input$ar_pop]], na.rm = TRUE),
        AttackRate = (sum(df[[input$ar_case]], na.rm = TRUE) / sum(df[[input$ar_pop]], na.rm = TRUE)) * 100
      )
    }
    
    datatable(ar_data)
  })
  
  # Attack Rates with Confidence Intervals
  output$ar_detailed_table <- renderDT({
    req(input$ar_case, input$ar_pop)
    df <- data()
    
    if (input$ar_group != "None") {
      ar_data <- df %>%
        group_by(.data[[input$ar_group]]) %>%
        summarise(
          Cases = sum(.data[[input$ar_case]], na.rm = TRUE),
          Population = sum(.data[[input$ar_pop]], na.rm = TRUE),
          AttackRate = (Cases / Population) * 100,
          LowerCI = prop.test(Cases, Population)$conf.int[1] * 100,
          UpperCI = prop.test(Cases, Population)$conf.int[2] * 100,
          .groups = "drop"
        )
    } else {
      total_cases <- sum(df[[input$ar_case]], na.rm = TRUE)
      total_pop <- sum(df[[input$ar_pop]], na.rm = TRUE)
      prop_test <- prop.test(total_cases, total_pop)
      
      ar_data <- data.frame(
        Cases = total_cases,
        Population = total_pop,
        AttackRate = (total_cases / total_pop) * 100,
        LowerCI = prop_test$conf.int[1] * 100,
        UpperCI = prop_test$conf.int[2] * 100
      )
    }
    
    datatable(ar_data)
  })
  
  # Download Handler for Attack Rate Table
  output$download_ar <- downloadHandler(
    filename = function() {
      "attack_rates.docx"
    },
    content = function(file) {
      df <- data()
      
      if (input$ar_format == "Table") {
        if (input$ar_group != "None") {
          ar_data <- df %>%
            group_by(.data[[input$ar_group]]) %>%
            summarise(
              Cases = sum(.data[[input$ar_case]], na.rm = TRUE),
              Population = sum(.data[[input$ar_pop]], na.rm = TRUE),
              AttackRate = (Cases / Population) * 100,
              .groups = "drop"
            )
        } else {
          ar_data <- data.frame(
            Cases = sum(df[[input$ar_case]], na.rm = TRUE),
            Population = sum(df[[input$ar_pop]], na.rm = TRUE),
            AttackRate = (sum(df[[input$ar_case]], na.rm = TRUE) / sum(df[[input$ar_pop]], na.rm = TRUE)) * 100
          )
        }
      } else {
        if (input$ar_group != "None") {
          ar_data <- df %>%
            group_by(.data[[input$ar_group]]) %>%
            summarise(
              Cases = sum(.data[[input$ar_case]], na.rm = TRUE),
              Population = sum(.data[[input$ar_pop]], na.rm = TRUE),
              AttackRate = (Cases / Population) * 100,
              LowerCI = prop.test(Cases, Population)$conf.int[1] * 100,
              UpperCI = prop.test(Cases, Population)$conf.int[2] * 100,
              .groups = "drop"
            )
        } else {
          total_cases <- sum(df[[input$ar_case]], na.rm = TRUE)
          total_pop <- sum(df[[input$ar_pop]], na.rm = TRUE)
          prop_test <- prop.test(total_cases, total_pop)
          
          ar_data <- data.frame(
            Cases = total_cases,
            Population = total_pop,
            AttackRate = (total_cases / total_pop) * 100,
            LowerCI = prop_test$conf.int[1] * 100,
            UpperCI = prop_test$conf.int[2] * 100
          )
        }
      }
      
      ft <- flextable(ar_data)
      doc <- read_docx() %>%
        body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  
  # Mortality Rates
  output$mort_table <- renderDT({
    req(input$mort_case, input$mort_pop)
    df <- data()
    
    if(input$mort_measure == "Mortality Rate") {
      total_cases <- sum(df[[input$mort_case]], na.rm = TRUE)
      total_pop <- sum(df[[input$mort_pop]], na.rm = TRUE)
      
      rate <- (total_cases / total_pop) * 1000  # per 1000 population
      
      result <- data.frame(
        Measure = "Mortality Rate",
        Deaths = total_cases,
        Population = total_pop,
        Rate = rate,
        Units = "per 1000 population"
      )
    } else {
      # Case Fatality Rate
      total_cases <- sum(df[[input$mort_case]], na.rm = TRUE)
      total_cases_all <- nrow(df)
      
      cfr <- (total_cases / total_cases_all) * 100  # percentage
      
      result <- data.frame(
        Measure = "Case Fatality Rate",
        Deaths = total_cases,
        Cases = total_cases_all,
        Rate = cfr,
        Units = "percent"
      )
    }
    
    datatable(result)
  })
  
  output$download_mort <- downloadHandler(
    filename = function() {
      "mortality_rates.docx"
    },
    content = function(file) {
      df <- data()
      
      if(input$mort_measure == "Mortality Rate") {
        total_cases <- sum(df[[input$mort_case]], na.rm = TRUE)
        total_pop <- sum(df[[input$mort_pop]], na.rm = TRUE)
        
        rate <- (total_cases / total_pop) * 1000
        
        result <- data.frame(
          Measure = "Mortality Rate",
          Deaths = total_cases,
          Population = total_pop,
          Rate = rate,
          Units = "per 1000 population"
        )
      } else {
        total_cases <- sum(df[[input$mort_case]], na.rm = TRUE)
        total_cases_all <- nrow(df)
        
        cfr <- (total_cases / total_cases_all) * 100
        
        result <- data.frame(
          Measure = "Case Fatality Rate",
          Deaths = total_cases,
          Cases = total_cases_all,
          Rate = cfr,
          Units = "percent"
        )
      }
      
      ft <- flextable(result)
      doc <- read_docx() %>% 
        body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # Life table analysis
  output$lifetable_results <- renderDT({
    req(input$lt_time, input$lt_event)
    df <- data()
    
    # Parse time breaks
    breaks <- as.numeric(unlist(strsplit(input$lt_breaks, ",")))
    
    # Create survival object
    time <- as.numeric(df[[input$lt_time]])
    event <- as.numeric(df[[input$lt_event]])
    surv_obj <- Surv(time = time, event = event)
    
    # Create life table
    lt <- survfit(surv_obj ~ 1)
    lt_summary <- summary(lt, times = breaks, extend = TRUE)
    
    # Create life table dataframe
    lifetable <- data.frame(
      Time = lt_summary$time,
      AtRisk = lt_summary$n.risk,
      Events = lt_summary$n.event,
      Survival = lt_summary$surv,
      Lower_CI = lt_summary$lower,
      Upper_CI = lt_summary$upper
    )
    
    datatable(lifetable, options = list(scrollX = TRUE, pageLength = 10))
  })
  
  output$download_lifetable <- downloadHandler(
    filename = function() {
      paste("life_table_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      df <- data()
      breaks <- as.numeric(unlist(strsplit(input$lt_breaks, ",")))
      time <- as.numeric(df[[input$lt_time]])
      event <- as.numeric(df[[input$lt_event]])
      surv_obj <- Surv(time = time, event = event)
      lt <- survfit(surv_obj ~ 1)
      lt_summary <- summary(lt, times = breaks, extend = TRUE)
      
      lifetable <- data.frame(
        Time = lt_summary$time,
        AtRisk = lt_summary$n.risk,
        Events = lt_summary$n.event,
        Survival = lt_summary$surv,
        Lower_CI = lt_summary$lower,
        Upper_CI = lt_summary$upper
      )
      
      ft <- flextable(lifetable) %>%
        set_caption(input$lt_title) %>%
        theme_zebra()
      
      doc <- read_docx() %>% 
        body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # Risk Ratios
  output$rr_results <- renderPrint({
    req(input$rr_outcome, input$rr_exposure)
    df <- data()
    
    # Create 2x2 table
    tab <- table(df[[input$rr_exposure]], df[[input$rr_outcome]])
    
    # Calculate risk ratio with confidence intervals
    rr <- riskratio(tab, conf.level = input$rr_conf)
    
    print(rr)
  })
  
  output$rr_table <- renderDT({
    req(input$rr_outcome, input$rr_exposure)
    df <- data()
    
    tab <- table(df[[input$rr_exposure]], df[[input$rr_outcome]])
    rr <- riskratio(tab, conf.level = input$rr_conf)
    
    # Create summary table
    rr_table <- data.frame(
      Measure = c("Risk Ratio", "Lower CI", "Upper CI"),
      Value = c(rr$measure[2,1], rr$measure[2,2], rr$measure[2,3])
    )
    
    datatable(rr_table)
  })
  
  output$download_rr <- downloadHandler(
    filename = function() {
      "risk_ratio.docx"
    },
    content = function(file) {
      df <- data()
      
      tab <- table(df[[input$rr_exposure]], df[[input$rr_outcome]])
      rr <- riskratio(tab, conf.level = input$rr_conf)
      
      rr_table <- data.frame(
        Measure = c("Risk Ratio", "Lower CI", "Upper CI"),
        Value = c(rr$measure[2,1], rr$measure[2,2], rr$measure[2,3])
      )
      
      ft <- flextable(rr_table)
      doc <- read_docx() %>% 
        body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # Odds Ratios
  output$or_results <- renderPrint({
    req(input$or_outcome, input$or_exposure)
    df <- data()
    
    # Create 2x2 table
    tab <- table(df[[input$or_exposure]], df[[input$or_outcome]])
    
    # Calculate odds ratio with confidence intervals
    or <- oddsratio(tab, conf.level = input$or_conf)
    
    print(or)
  })
  
  output$or_table <- renderDT({
    req(input$or_outcome, input$or_exposure)
    df <- data()
    
    tab <- table(df[[input$or_exposure]], df[[input$or_outcome]])
    or <- oddsratio(tab, conf.level = input$or_conf)
    
    # Create summary table
    or_table <- data.frame(
      Measure = c("Odds Ratio", "Lower CI", "Upper CI"),
      Value = c(or$measure[2,1], or$measure[2,2], or$measure[2,3])
    )
    
    datatable(or_table)
  })
  
  output$download_or <- downloadHandler(
    filename = function() {
      "odds_ratio.docx"
    },
    content = function(file) {
      df <- data()
      
      tab <- table(df[[input$or_exposure]], df[[input$or_outcome]])
      or <- oddsratio(tab, conf.level = input$or_conf)
      
      or_table <- data.frame(
        Measure = c("Odds Ratio", "Lower CI", "Upper CI"),
        Value = c(or$measure[2,1], or$measure[2,2], or$measure[2,3])
      )
      
      ft <- flextable(or_table)
      doc <- read_docx() %>% 
        body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # Attributable Risk
  output$ar_results <- renderPrint({
    req(input$ar_outcome, input$ar_exposure)
    df <- data()
    
    # Create 2x2 table
    tab <- table(df[[input$ar_exposure]], df[[input$ar_outcome]])
    
    # Calculate attributable risk
    ar <- riskratio(tab, conf.level = input$ar_conf)
    
    print(ar)
  })
  
  output$ar_table <- renderDT({
    req(input$ar_outcome, input$ar_exposure)
    df <- data()
    
    tab <- table(df[[input$ar_exposure]], df[[input$ar_outcome]])
    ar <- riskratio(tab, conf.level = input$ar_conf)
    
    # Create summary table
    ar_table <- data.frame(
      Measure = c("Attributable Risk", "Lower CI", "Upper CI"),
      Value = c(ar$measure[1,1], ar$measure[1,2], ar$measure[1,3])
    )
    
    if(input$ar_paf) {
      # Calculate population attributable fraction
      paf <- (ar$measure[1,1] - 1) / ar$measure[1,1]
      ar_table <- rbind(ar_table, data.frame(Measure = "Population Attributable Fraction", Value = paf))
    }
    
    datatable(ar_table)
  })
  
  output$download_arisk <- downloadHandler(
    filename = function() {
      "attributable_risk.docx"
    },
    content = function(file) {
      df <- data()
      
      tab <- table(df[[input$ar_exposure]], df[[input$ar_outcome]])
      ar <- riskratio(tab, conf.level = input$ar_conf)
      
      ar_table <- data.frame(
        Measure = c("Attributable Risk", "Lower CI", "Upper CI"),
        Value = c(ar$measure[1,1], ar$measure[1,2], ar$measure[1,3])
      )
      
      if(input$ar_paf) {
        paf <- (ar$measure[1,1] - 1) / ar$measure[1,1]
        ar_table <- rbind(ar_table, data.frame(Measure = "Population Attributable Fraction", Value = paf))
      }
      
      ft <- flextable(ar_table)
      doc <- read_docx() %>% 
        body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # Poisson Regression Reference UI
  output$poisson_ref_ui <- renderUI({
    req(input$poisson_vars)
    df <- data()
    
    ref_inputs <- lapply(input$poisson_vars, function(var) {
      if (is.factor(df[[var]])) {
        selectInput(paste0("poisson_ref_", var),
                    paste("Reference level for", var),
                    choices = levels(df[[var]]))
      } else if (is.character(df[[var]])) {
        selectInput(paste0("poisson_ref_", var),
                    paste("Reference level for", var),
                    choices = unique(df[[var]]))
      }
    })
    
    do.call(tagList, ref_inputs)
  })
  
  # Poisson Regression Summary Output
  output$poisson_summary <- renderPrint({
    req(input$poisson_outcome, input$poisson_vars)
    df <- data()
    
    # Set reference levels
    for (var in input$poisson_vars) {
      ref_input <- input[[paste0("poisson_ref_", var)]]
      if (!is.null(ref_input) && (is.factor(df[[var]]) || is.character(df[[var]]))) {
        df[[var]] <- relevel(as.factor(df[[var]]), ref = ref_input)
      }
    }
    
    # Create formula
    formula <- as.formula(paste(input$poisson_outcome, "~", paste(input$poisson_vars, collapse = "+")))
    
    # Fit model
    if (!is.null(input$poisson_offset) && input$poisson_offset != "") {
      model <- glm(formula, family = poisson(link = "log"),
                   data = df, offset = log(df[[input$poisson_offset]]))
    } else {
      model <- glm(formula, family = poisson(link = "log"), data = df)
    }
    
    # Print summary
    summary(model)
  })
  
  # Poisson Regression Table Output
  output$poisson_table <- renderDT({
    req(input$poisson_outcome, input$poisson_vars)
    df <- data()
    
    # Set reference levels
    for (var in input$poisson_vars) {
      ref_input <- input[[paste0("poisson_ref_", var)]]
      if (!is.null(ref_input) && (is.factor(df[[var]]) || is.character(df[[var]]))) {
        df[[var]] <- relevel(as.factor(df[[var]]), ref = ref_input)
      }
    }
    
    # Create formula
    formula <- as.formula(paste(input$poisson_outcome, "~", paste(input$poisson_vars, collapse = "+")))
    
    # Fit model
    if (!is.null(input$poisson_offset) && input$poisson_offset != "") {
      model <- glm(formula, family = poisson(link = "log"),
                   data = df, offset = log(df[[input$poisson_offset]]))
    } else {
      model <- glm(formula, family = poisson(link = "log"), data = df)
    }
    
    # Create tidy table
    if (input$poisson_irr) {
      poisson_table <- broom::tidy(model, exponentiate = TRUE, conf.int = TRUE, conf.level = input$poisson_conf)
    } else {
      poisson_table <- broom::tidy(model, conf.int = TRUE, conf.level = input$poisson_conf)
    }
    
    datatable(poisson_table)
  })
  
  # Poisson Regression Download
  output$download_poisson <- downloadHandler(
    filename = function() {
      "poisson_regression.docx"
    },
    content = function(file) {
      df <- data()
      
      # Set reference levels
      for (var in input$poisson_vars) {
        ref_input <- input[[paste0("poisson_ref_", var)]]
        if (!is.null(ref_input) && (is.factor(df[[var]]) || is.character(df[[var]]))) {
          df[[var]] <- relevel(as.factor(df[[var]]), ref = ref_input)
        }
      }
      
      # Create formula
      formula <- as.formula(paste(input$poisson_outcome, "~", paste(input$poisson_vars, collapse = "+")))
      
      # Fit model
      if (!is.null(input$poisson_offset) && input$poisson_offset != "") {
        model <- glm(formula, family = poisson(link = "log"),
                     data = df, offset = log(df[[input$poisson_offset]]))
      } else {
        model <- glm(formula, family = poisson(link = "log"), data = df)
      }
      
      # Create tidy table
      if (input$poisson_irr) {
        poisson_table <- broom::tidy(model, exponentiate = TRUE, conf.int = TRUE, conf.level = input$poisson_conf)
      } else {
        poisson_table <- broom::tidy(model, conf.int = TRUE, conf.level = input$poisson_conf)
      }
      
      ft <- flextable::flextable(poisson_table)
      doc <- officer::read_docx() %>%
        body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # Survival Analysis (Kaplan-Meier)
  output$surv_summary <- renderPrint({
    req(input$surv_time, input$surv_event)
    df <- data()
    
    # Create survival object
    surv_obj <- Surv(df[[input$surv_time]], df[[input$surv_event]])
    
    # Fit survival model
    if(input$surv_group != "None") {
      surv_fit <- survfit(surv_obj ~ .data[[input$surv_group]], data = df)
    } else {
      surv_fit <- survfit(surv_obj ~ 1, data = df)
    }
    
    # Print summary
    print(surv_fit)
    
    # Show median survival if requested
    if(input$surv_median) {
      cat("\nMedian Survival:\n")
      print(surv_median(surv_fit))
    }
    
    # Show mean survival if requested
    if(input$surv_mean) {
      cat("\nMean Survival:\n")
      print(surv_mean(surv_fit))
    }
    
    # Show hazard ratio if requested and groups exist
    if(input$surv_hazard && input$surv_group != "None") {
      cat("\nHazard Ratio:\n")
      cox_model <- coxph(surv_obj ~ .data[[input$surv_group]], data = df)
      print(summary(cox_model))
    }
    
    # Test proportional hazards if requested
    if(input$surv_ph && input$surv_group != "None") {
      cat("\nProportional Hazards Test:\n")
      print(cox.zph(coxph(surv_obj ~ .data[[input$surv_group]], data = df)))
    }
  })
  
  output$surv_plot <- renderPlot({
    req(input$surv_time, input$surv_event)
    df <- data()
    
    # Create survival object
    surv_obj <- Surv(df[[input$surv_time]], df[[input$surv_event]])
    
    # Fit survival model
    if(input$surv_group != "None") {
      surv_fit <- survfit(surv_obj ~ .data[[input$surv_group]], data = df)
      
      # Customize legend labels if provided
      legend_labels <- unlist(strsplit(input$surv_legend_labels, ","))
      if(length(legend_labels) > 0 && length(legend_labels) == length(levels(as.factor(df[[input$surv_group]])))) {
        names(legend_labels) <- levels(as.factor(df[[input$surv_group]]))
      } else {
        legend_labels <- NULL
      }
    } else {
      surv_fit <- survfit(surv_obj ~ 1, data = df)
      legend_labels <- NULL
    }
    
    # Create plot
    p <- ggsurvplot(
      surv_fit,
      data = df,
      risk.table = TRUE,
      pval = TRUE,
      conf.int = TRUE,
      palette = c(input$surv_color1, input$surv_color2),
      title = input$surv_title,
      xlab = input$surv_xlab,
      ylab = input$surv_ylab,
      legend.title = input$surv_legend_title,
      legend.labs = legend_labels,
      break.time.by = input$surv_break_time,
      ggtheme = theme_minimal(),
      font.title = c(input$surv_title_size, "bold"),
      font.x = c(input$surv_x_size, "plain"),
      font.y = c(input$surv_y_size, "plain"),
      font.tickslab = c(12, "plain")
    )
    
    print(p)
  })
  
  output$download_surv <- downloadHandler(
    filename = function() {
      "survival_plot.png"
    },
    content = function(file) {
      df <- data()
      
      surv_obj <- Surv(df[[input$surv_time]], df[[input$surv_event]])
      
      if(input$surv_group != "None") {
        surv_fit <- survfit(surv_obj ~ .data[[input$surv_group]], data = df)
        
        legend_labels <- unlist(strsplit(input$surv_legend_labels, ","))
        if(length(legend_labels) > 0 && length(legend_labels) == length(levels(as.factor(df[[input$surv_group]])))) {
          names(legend_labels) <- levels(as.factor(df[[input$surv_group]]))
        } else {
          legend_labels <- NULL
        }
      } else {
        surv_fit <- survfit(surv_obj ~ 1, data = df)
        legend_labels <- NULL
      }
      
      p <- ggsurvplot(
        surv_fit,
        data = df,
        risk.table = TRUE,
        pval = TRUE,
        conf.int = TRUE,
        palette = c(input$surv_color1, input$surv_color2),
        title = input$surv_title,
        xlab = input$surv_xlab,
        ylab = input$surv_ylab,
        legend.title = input$surv_legend_title,
        legend.labs = legend_labels,
        break.time.by = input$surv_break_time,
        ggtheme = theme_minimal(),
        font.title = c(input$surv_title_size, "bold"),
        font.x = c(input$surv_x_size, "plain"),
        font.y = c(input$surv_y_size, "plain"),
        font.tickslab = c(12, "plain")
      )
      
      ggsave(file, plot = print(p), device = "png", width = 10, height = 8)
    }
  )
  
  # Log-rank test
  output$logrank_results <- renderPrint({
    req(input$lr_time, input$lr_event, input$lr_group)
    df <- data()
    
    surv_obj <- tryCatch({
      create_surv_obj(input$lr_time, input$lr_event, df)
    }, error = function(e) {
      # Fallback to simple creation if specialized function fails
      time <- df[[input$lr_time]]
      event <- df[[input$lr_event]]
      Surv(time = time, event = as.numeric(event))
    })
    
    fit <- survdiff(surv_obj ~ df[[input$lr_group]])
    
    cat("Log-Rank Test Results:\n")
    print(fit)
  })
  
  output$download_logrank <- downloadHandler(
    filename = "logrank_test.docx",
    content = function(file) {
      df <- data()
      
      surv_obj <- tryCatch({
        create_surv_obj(input$lr_time, input$lr_event, df)
      }, error = function(e) {
        # Fallback to simple creation if specialized function fails
        time <- df[[input$lr_time]]
        event <- df[[input$lr_event]]
        Surv(time = time, event = as.numeric(event))
      })
      
      fit <- survdiff(surv_obj ~ df[[input$lr_group]])
      
      # Convert to flextable
      res <- capture.output(fit)
      ft <- flextable(data.frame(Results = res))
      doc <- read_docx() %>% 
        body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  
  # Cox Regression
  output$cox_ref_ui <- renderUI({
    req(input$cox_vars)
    df <- data()
    
    ref_inputs <- lapply(input$cox_vars, function(var) {
      if(is.factor(df[[var]])) {
        selectInput(paste0("cox_ref_", var), 
                    paste("Reference level for", var), 
                    choices = levels(df[[var]]))
      } else if(is.character(df[[var]])) {
        selectInput(paste0("cox_ref_", var), 
                    paste("Reference level for", var), 
                    choices = unique(df[[var]]))
      }
    })
    
    do.call(tagList, ref_inputs)
  })
  
  output$cox_summary <- renderPrint({
    req(input$cox_time, input$cox_event, input$cox_vars)
    df <- data()
    
    # Set reference levels
    for(var in input$cox_vars) {
      ref_input <- input[[paste0("cox_ref_", var)]]
      if(!is.null(ref_input) && (is.factor(df[[var]]) || is.character(df[[var]]))) {
        df[[var]] <- relevel(as.factor(df[[var]]), ref = ref_input)
      }
    }
    
    # Create survival object
    surv_obj <- Surv(df[[input$cox_time]], df[[input$cox_event]])
    
    # Create formula
    formula <- as.formula(paste("surv_obj ~", paste(input$cox_vars, collapse = "+")))
    
    # Fit Cox model
    model <- coxph(formula, data = df)
    
    # Print summary
    summary(model)
  })
  
  output$cox_table <- renderDT({
    req(input$cox_time, input$cox_event, input$cox_vars)
    df <- data()
    
    # Set reference levels
    for(var in input$cox_vars) {
      ref_input <- input[[paste0("cox_ref_", var)]]
      if(!is.null(ref_input) && (is.factor(df[[var]]) || is.character(df[[var]]))) {
        df[[var]] <- relevel(as.factor(df[[var]]), ref = ref_input)
      }
    }
    
    # Create survival object
    surv_obj <- Surv(df[[input$cox_time]], df[[input$cox_event]])
    
    # Create formula
    formula <- as.formula(paste("surv_obj ~", paste(input$cox_vars, collapse = "+")))
    
    # Fit Cox model
    model <- coxph(formula, data = df)
    
    # Create tidy table
    cox_table <- tidy(model, conf.int = TRUE, conf.level = input$cox_conf, exponentiate = TRUE)
    
    datatable(cox_table)
  })
  
  output$download_cox <- downloadHandler(
    filename = function() {
      "cox_regression.docx"
    },
    content = function(file) {
      df <- data()
      
      # Set reference levels
      for(var in input$cox_vars) {
        ref_input <- input[[paste0("cox_ref_", var)]]
        if(!is.null(ref_input) && (is.factor(df[[var]]) || is.character(df[[var]]))) {
          df[[var]] <- relevel(as.factor(df[[var]]), ref = ref_input)
        }
      }
      
      surv_obj <- Surv(df[[input$cox_time]], df[[input$cox_event]])
      formula <- as.formula(paste("surv_obj ~", paste(input$cox_vars, collapse = "+")))
      model <- coxph(formula, data = df)
      
      cox_table <- tidy(model, conf.int = TRUE, conf.level = input$cox_conf, exponentiate = TRUE)
      
      ft <- flextable(cox_table)
      doc <- read_docx() %>% 
        body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # Linear regression with reference category selection
  output$linear_ref_ui <- renderUI({
    req(input$linear_vars)
    df <- data()
    
    ref_ui <- tagList()
    
    for (var in input$linear_vars) {
      if (is.factor(df[[var]]) || is.character(df[[var]])) {
        levels <- if (is.factor(df[[var]])) levels(df[[var]]) else unique(df[[var]])
        ref_ui <- tagAppendChild(
          ref_ui,
          selectInput(paste0("linear_ref_", var), 
                      paste("Reference category for", var), 
                      choices = levels)
        )
      }
    }
    
    ref_ui
  })
  
  output$linear_summary <- renderPrint({
    req(input$linear_outcome, input$linear_vars)
    df <- data()
    
    # Set reference categories for factors
    for (var in input$linear_vars) {
      if (is.factor(df[[var]])) {
        ref <- input[[paste0("linear_ref_", var)]]
        if (!is.null(ref)) {
          df[[var]] <- relevel(df[[var]], ref = ref)
        }
      } else if (is.character(df[[var]])) {
        ref <- input[[paste0("linear_ref_", var)]]
        if (!is.null(ref)) {
          df[[var]] <- factor(df[[var]])
          df[[var]] <- relevel(df[[var]], ref = ref)
        }
      }
    }
    
    # Create formula
    formula <- as.formula(paste(input$linear_outcome, "~", paste(input$linear_vars, collapse = " + ")))
    
    # Fit linear model
    linear_fit <- lm(formula, data = df)
    
    cat("Linear Regression Results:\n\n")
    print(summary(linear_fit))
    
    if (input$linear_std) {
      cat("\nStandardized Coefficients:\n")
      # Calculate standardized coefficients
      std_coef <- data.frame(
        term = names(coef(linear_fit)),
        estimate = lm.beta::lm.beta(linear_fit)$standardized.coefficients
      )
      print(std_coef)
    }
  })
  
  output$linear_table <- renderDT({
    req(input$linear_outcome, input$linear_vars)
    df <- data()
    
    # Set reference categories for factors
    for (var in input$linear_vars) {
      if (is.factor(df[[var]])) {
        ref <- input[[paste0("linear_ref_", var)]]
        if (!is.null(ref)) {
          df[[var]] <- relevel(df[[var]], ref = ref)
        }
      } else if (is.character(df[[var]])) {
        ref <- input[[paste0("linear_ref_", var)]]
        if (!is.null(ref)) {
          df[[var]] <- factor(df[[var]])
          df[[var]] <- relevel(df[[var]], ref = ref)
        }
      }
    }
    
    # Create formula
    formula <- as.formula(paste(input$linear_outcome, "~", paste(input$linear_vars, collapse = " + ")))
    
    # Fit linear model
    linear_fit <- lm(formula, data = df)
    
    # Create results table
    results <- broom::tidy(linear_fit, conf.int = TRUE) %>%
      select(term, estimate, conf.low, conf.high, p.value) %>%
      rename(Coefficient = estimate, `Lower CI` = conf.low, `Upper CI` = conf.high, `p-value` = p.value)
    
    if (input$linear_std) {
      # Add standardized coefficients
      std_coef <- data.frame(
        term = names(coef(linear_fit)),
        std_estimate = lm.beta::lm.beta(linear_fit)$standardized.coefficients
      )
      results <- results %>%
        left_join(std_coef, by = "term") %>%
        rename(`Standardized Coefficient` = std_estimate)
    }
    
    datatable(results, options = list(pageLength = 10))
  })
  
  output$download_linear <- downloadHandler(
    filename = "linear_regression.docx",
    content = function(file) {
      df <- data()
      
      # Set reference categories for factors
      for (var in input$linear_vars) {
        if (is.factor(df[[var]]) || is.character(df[[var]])) {
          ref <- input[[paste0("linear_ref_", var)]]
          if (!is.null(ref)) {
            df[[var]] <- factor(df[[var]])
            df[[var]] <- relevel(df[[var]], ref = ref)
          }
        }
      }
      
      formula <- as.formula(paste(input$linear_outcome, "~", paste(input$linear_vars, collapse = " + ")))
      linear_fit <- lm(formula, data = df)
      
      # Create results table
      results <- broom::tidy(linear_fit, conf.int = TRUE) %>%
        select(term, estimate, conf.low, conf.high, p.value) %>%
        rename(Coefficient = estimate, `Lower CI` = conf.low, `Upper CI` = conf.high, `p-value` = p.value)
      
      if (input$linear_std) {
        # Add standardized coefficients
        std_coef <- data.frame(
          term = names(coef(linear_fit)),
          std_estimate = lm.beta::lm.beta(linear_fit)$standardized.coefficients
        )
        results <- results %>%
          left_join(std_coef, by = "term") %>%
          rename(`Standardized Coefficient` = std_estimate)
      }
      
      ft <- flextable(results) %>%
        set_caption("Linear Regression Results") %>%
        theme_zebra()
      
      doc <- read_docx() %>% 
        body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # Logistic Regression
  output$logistic_ref_ui <- renderUI({
    req(input$logistic_vars)
    df <- data()
    
    ref_inputs <- lapply(input$logistic_vars, function(var) {
      if(is.factor(df[[var]])) {
        selectInput(paste0("logistic_ref_", var), 
                    paste("Reference level for", var), 
                    choices = levels(df[[var]]))
      } else if(is.character(df[[var]])) {
        selectInput(paste0("logistic_ref_", var), 
                    paste("Reference level for", var), 
                    choices = unique(df[[var]]))
      }
    })
    
    do.call(tagList, ref_inputs)
  })
  
  output$logistic_summary <- renderPrint({
    req(input$logistic_outcome, input$logistic_vars)
    df <- data()
    
    # Set reference levels
    for(var in input$logistic_vars) {
      ref_input <- input[[paste0("logistic_ref_", var)]]
      if(!is.null(ref_input) && (is.factor(df[[var]]) || is.character(df[[var]]))) {
        df[[var]] <- relevel(as.factor(df[[var]]), ref = ref_input)
      }
    }
    
    # Create formula
    formula <- as.formula(paste(input$logistic_outcome, "~", paste(input$logistic_vars, collapse = "+")))
    
    # Fit appropriate model based on type
    if(input$logistic_type == "binary") {
      model <- glm(formula, family = binomial(link = "logit"), data = df)
    } else if(input$logistic_type == "ordinal") {
      req(requireNamespace("MASS", quietly = TRUE))
      model <- MASS::polr(formula, data = df, Hess = TRUE)
    } else {
      req(requireNamespace("nnet", quietly = TRUE))
      model <- nnet::multinom(formula, data = df)
    }
    
    # Print summary
    print(summary(model))
    
    # Show odds ratios for binary logistic regression
    if(input$logistic_or && input$logistic_type == "binary") {
      cat("\nOdds Ratios:\n")
      or <- exp(coef(model))
      ci <- exp(confint(model))
      or_table <- data.frame(
        OR = or,
        LowerCI = ci[,1],
        UpperCI = ci[,2]
      )
      print(or_table)
    }
  })
  
  output$logistic_table <- renderDT({
    req(input$logistic_outcome, input$logistic_vars)
    df <- data()
    
    # Set reference levels
    for(var in input$logistic_vars) {
      ref_input <- input[[paste0("logistic_ref_", var)]]
      if(!is.null(ref_input) && (is.factor(df[[var]]) || is.character(df[[var]]))) {
        df[[var]] <- relevel(as.factor(df[[var]]), ref = ref_input)
      }
    }
    
    # Create formula
    formula <- as.formula(paste(input$logistic_outcome, "~", paste(input$logistic_vars, collapse = "+")))
    
    # Fit appropriate model
    if(input$logistic_type == "binary") {
      model <- glm(formula, family = binomial(link = "logit"), data = df)
      
      # Create tidy table
      if(input$logistic_or) {
        logistic_table <- tidy(model, conf.int = TRUE, exponentiate = TRUE)
      } else {
        logistic_table <- tidy(model, conf.int = TRUE)
      }
    } else if(input$logistic_type == "ordinal") {
      req(requireNamespace("MASS", quietly = TRUE))
      model <- MASS::polr(formula, data = df, Hess = TRUE)
      logistic_table <- as.data.frame(coef(summary(model)))
      logistic_table$p.value <- pnorm(abs(logistic_table[, "t value"]), lower.tail = FALSE) * 2
    } else {
      req(requireNamespace("nnet", quietly = TRUE))
      model <- nnet::multinom(formula, data = df)
      logistic_table <- as.data.frame(coef(summary(model)))
    }
    
    datatable(logistic_table)
  })
  
  output$download_logistic <- downloadHandler(
    filename = function() {
      "logistic_regression.docx"
    },
    content = function(file) {
      df <- data()
      
      # Set reference levels
      for(var in input$logistic_vars) {
        ref_input <- input[[paste0("logistic_ref_", var)]]
        if(!is.null(ref_input) && (is.factor(df[[var]]) || is.character(df[[var]]))) {
          df[[var]] <- relevel(as.factor(df[[var]]), ref = ref_input)
        }
      }
      
      formula <- as.formula(paste(input$logistic_outcome, "~", paste(input$logistic_vars, collapse = "+")))
      
      if(input$logistic_type == "binary") {
        model <- glm(formula, family = binomial(link = "logit"), data = df)
        
        if(input$logistic_or) {
          logistic_table <- tidy(model, conf.int = TRUE, exponentiate = TRUE)
        } else {
          logistic_table <- tidy(model, conf.int = TRUE)
        }
      } else if(input$logistic_type == "ordinal") {
        req(requireNamespace("MASS", quietly = TRUE))
        model <- MASS::polr(formula, data = df, Hess = TRUE)
        logistic_table <- as.data.frame(coef(summary(model)))
        logistic_table$p.value <- pnorm(abs(logistic_table[, "t value"]), lower.tail = FALSE) * 2
      } else {
        req(requireNamespace("nnet", quietly = TRUE))
        model <- nnet::multinom(formula, data = df)
        logistic_table <- as.data.frame(coef(summary(model)))
      }
      
      ft <- flextable(logistic_table)
      doc <- read_docx() %>% 
        body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # T-tests/Mann-Whitney
  output$ttest_results <- renderPrint({
    req(input$ttest_var)
    df <- data()
    
    if(input$ttest_type == "one.sample") {
      # One sample t-test
      t_test <- t.test(df[[input$ttest_var]], mu = input$ttest_mu)
      print(t_test)
    } else if(input$ttest_type == "two.sample") {
      # Two sample t-test
      req(input$ttest_group)
      t_test <- t.test(df[[input$ttest_var]] ~ df[[input$ttest_group]], 
                       var.equal = input$ttest_var_equal)
      print(t_test)
    } else if(input$ttest_type == "paired") {
      # Paired t-test
      req(input$ttest_group)
      t_test <- t.test(df[[input$ttest_var]], df[[input$ttest_group]], paired = TRUE)
      print(t_test)
    } else {
      # Mann-Whitney U test
      req(input$ttest_group)
      mw_test <- wilcox.test(df[[input$ttest_var]] ~ df[[input$ttest_group]])
      print(mw_test)
    }
  })
  
  output$download_ttest <- downloadHandler(
    filename = function() {
      "t_test_results.docx"
    },
    content = function(file) {
      df <- data()
      
      if(input$ttest_type == "one.sample") {
        t_test <- t.test(df[[input$ttest_var]], mu = input$ttest_mu)
        results <- data.frame(
          Test = "One Sample t-test",
          Statistic = t_test$statistic,
          DF = t_test$parameter,
          PValue = t_test$p.value,
          Estimate = t_test$estimate,
          ConfLow = t_test$conf.int[1],
          ConfHigh = t_test$conf.int[2]
        )
      } else if(input$ttest_type == "two.sample") {
        t_test <- t.test(df[[input$ttest_var]] ~ df[[input$ttest_group]], 
                         var.equal = input$ttest_var_equal)
        results <- data.frame(
          Test = "Two Sample t-test",
          Statistic = t_test$statistic,
          DF = t_test$parameter,
          PValue = t_test$p.value,
          Mean1 = t_test$estimate[1],
          Mean2 = t_test$estimate[2],
          ConfLow = t_test$conf.int[1],
          ConfHigh = t_test$conf.int[2]
        )
      } else if(input$ttest_type == "paired") {
        t_test <- t.test(df[[input$ttest_var]], df[[input$ttest_group]], paired = TRUE)
        results <- data.frame(
          Test = "Paired t-test",
          Statistic = t_test$statistic,
          DF = t_test$parameter,
          PValue = t_test$p.value,
          MeanDiff = t_test$estimate,
          ConfLow = t_test$conf.int[1],
          ConfHigh = t_test$conf.int[2]
        )
      } else {
        mw_test <- wilcox.test(df[[input$ttest_var]] ~ df[[input$ttest_group]])
        results <- data.frame(
          Test = "Mann-Whitney U Test",
          Statistic = mw_test$statistic,
          PValue = mw_test$p.value
        )
      }
      
      ft <- flextable(results)
      doc <- read_docx() %>% 
        body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # ANOVA/Kruskal-Wallis
  output$anova_results <- renderPrint({
    req(input$anova_var, input$anova_group)
    df <- data()
    
    if(input$anova_type == "anova") {
      # One-way ANOVA
      anova_model <- aov(df[[input$anova_var]] ~ df[[input$anova_group]])
      print(summary(anova_model))
    } else {
      # Kruskal-Wallis test
      kw_test <- kruskal.test(df[[input$anova_var]] ~ df[[input$anova_group]])
      print(kw_test)
    }
  })
  
  output$posthoc_results <- renderPrint({
    req(input$anova_var, input$anova_group, input$anova_posthoc)
    df <- data()
    
    if(input$anova_type == "anova") {
      # Tukey HSD post-hoc
      anova_model <- aov(df[[input$anova_var]] ~ df[[input$anova_group]])
      tukey <- TukeyHSD(anova_model)
      print(tukey)
    } else {
      # Dunn's test for Kruskal-Wallis
      req(requireNamespace("dunn.test", quietly = TRUE))
      dunn <- dunn.test::dunn.test(df[[input$anova_var]], df[[input$anova_group]], 
                                   method = "bonferroni")
      print(dunn)
    }
  })
  
  output$download_anova <- downloadHandler(
    filename = function() {
      "anova_results.docx"
    },
    content = function(file) {
      df <- data()
      
      if(input$anova_type == "anova") {
        anova_model <- aov(df[[input$anova_var]] ~ df[[input$anova_group]])
        anova_table <- as.data.frame(summary(anova_model)[[1]])
        
        if(input$anova_posthoc) {
          tukey <- as.data.frame(TukeyHSD(anova_model)[[1]])
          tukey$Comparison <- rownames(tukey)
          rownames(tukey) <- NULL
        }
      } else {
        kw_test <- kruskal.test(df[[input$anova_var]] ~ df[[input$anova_group]])
        anova_table <- data.frame(
          Test = "Kruskal-Wallis",
          ChiSquared = kw_test$statistic,
          DF = kw_test$parameter,
          PValue = kw_test$p.value
        )
        
        if(input$anova_posthoc) {
          req(requireNamespace("dunn.test", quietly = TRUE))
          dunn <- dunn.test::dunn.test(df[[input$anova_var]], df[[input$anova_group]], 
                                       method = "bonferroni")
          tukey <- data.frame(
            Comparison = dunn$comparisons,
            Z = dunn$Z,
            PValue = dunn$P.adjusted
          )
        }
      }
      
      ft1 <- flextable(anova_table)
      
      if(input$anova_posthoc) {
        ft2 <- flextable(tukey)
        doc <- read_docx() %>% 
          body_add_flextable(ft1) %>% 
          body_add_par("Post-hoc Tests:", style = "heading 2") %>% 
          body_add_flextable(ft2)
      } else {
        doc <- read_docx() %>% 
          body_add_flextable(ft1)
      }
      
      print(doc, target = file)
    }
  )
  
  # Chi-square Test
  output$chisq_results <- renderPrint({
    req(input$chisq_var1, input$chisq_var2)
    df <- data()
    
    # Create contingency table
    tab <- table(df[[input$chisq_var1]], df[[input$chisq_var2]])
    
    # Perform chi-square test
    chisq_test <- chisq.test(tab)
    
    # Print results
    print(chisq_test)
    
    # Show expected counts if requested
    if(input$chisq_expected) {
      cat("\nExpected Counts:\n")
      print(chisq_test$expected)
    }
    
    # Show residuals if requested
    if(input$chisq_residuals) {
      cat("\nPearson Residuals:\n")
      print(chisq_test$residuals)
    }
  })
  
  output$chisq_table <- renderDT({
    req(input$chisq_var1, input$chisq_var2)
    df <- data()
    
    # Create contingency table
    tab <- as.data.frame.matrix(table(df[[input$chisq_var1]], df[[input$chisq_var2]]))
    
    datatable(tab)
  })
  
  output$download_chisq <- downloadHandler(
    filename = function() {
      "chi_square_test.docx"
    },
    content = function(file) {
      df <- data()
      
      tab <- table(df[[input$chisq_var1]], df[[input$chisq_var2]])
      chisq_test <- chisq.test(tab)
      
      results <- data.frame(
        Test = "Chi-square Test",
        Statistic = chisq_test$statistic,
        DF = chisq_test$parameter,
        PValue = chisq_test$p.value
      )
      
      ft1 <- flextable(results)
      ft2 <- flextable(as.data.frame.matrix(tab))
      
      doc <- read_docx() %>% 
        body_add_flextable(ft1) %>% 
        body_add_par("Contingency Table:", style = "heading 2") %>% 
        body_add_flextable(ft2)
      
      if(input$chisq_expected) {
        ft3 <- flextable(as.data.frame.matrix(chisq_test$expected))
        doc <- doc %>% 
          body_add_par("Expected Counts:", style = "heading 2") %>% 
          body_add_flextable(ft3)
      }
      
      if(input$chisq_residuals) {
        ft4 <- flextable(as.data.frame.matrix(chisq_test$residuals))
        doc <- doc %>% 
          body_add_par("Pearson Residuals:", style = "heading 2") %>% 
          body_add_flextable(ft4)
      }
      
      print(doc, target = file)
    }
  )
}

# Run the application
shinyApp(ui = ui, server = server)
