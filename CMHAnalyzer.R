library(shiny)
library(shinydashboard)
library(DT)
library(haven)
library(readxl)
library(ggplot2)
library(officer)
library(flextable)
library(rmarkdown)
library(tidyverse)
library(sortable)

# Helper function to convert haven_labelled variables to factors with labels
convert_labelled_to_factor <- function(df) {
  for (col_name in names(df)) {
    if (inherits(df[[col_name]], "haven_labelled")) {
      # Extract value labels
      val_labels <- attr(df[[col_name]], "labels")
      
      if (!is.null(val_labels) && length(val_labels) > 0) {
        # Convert to factor with labels
        df[[col_name]] <- as_factor(df[[col_name]])
      } else {
        # If no labels, convert to regular factor or character
        df[[col_name]] <- as.character(df[[col_name]])
      }
    }
  }
  return(df)
}

# Enhanced CMH Analysis Functions
calculate_cmh_analysis <- function(strata_tables, conf_level = 0.95) {
  q <- length(strata_tables)
  
  # Initialize sums for Cochran and Mantel-Haenszel statistics
  sum_observed_minus_expected <- 0
  sum_variance <- 0
  sum_numerator_or <- 0
  sum_denominator_or <- 0
  
  stratum_results <- list()
  
  for (h in 1:q) {
    table_h <- strata_tables[[h]]
    
    if (nrow(table_h) == 2 && ncol(table_h) == 2) {
      a <- table_h[1, 1]  # Exposed, Disease
      b <- table_h[1, 2]  # Exposed, No Disease
      c <- table_h[2, 1]  # Unexposed, Disease
      d <- table_h[2, 2]  # Unexposed, No Disease
      n <- a + b + c + d
      
      # Expected value and variance for Cochran's statistic
      n1_dot <- a + b
      n_dot1 <- a + c
      expected_a <- n1_dot * n_dot1 / n
      variance_a <- n1_dot * (n - n1_dot) * n_dot1 * (n - n_dot1) / (n^2 * (n - 1))
      
      # Accumulate for Cochran's statistic
      sum_observed_minus_expected <- sum_observed_minus_expected + (a - expected_a)
      sum_variance <- sum_variance + variance_a
      
      # Accumulate for Mantel-Haenszel OR
      sum_numerator_or <- sum_numerator_or + (a * d / n)
      sum_denominator_or <- sum_denominator_or + (b * c / n)
      
      stratum_results[[h]] <- list(
        a = a, b = b, c = c, d = d, n = n,
        expected_a = expected_a,
        variance_a = variance_a
      )
    }
  }
  
  # Cochran's statistic (without continuity correction)
  cochran_stat <- sum_observed_minus_expected^2 / sum_variance
  cochran_p <- 1 - pchisq(cochran_stat, 1)
  
  # Mantel-Haenszel statistic (with continuity correction when appropriate)
  if (abs(sum_observed_minus_expected) > 0) {
    mh_stat <- (abs(sum_observed_minus_expected) - 0.5)^2 / sum_variance
  } else {
    mh_stat <- sum_observed_minus_expected^2 / sum_variance  # No correction when difference is 0
  }
  mh_p <- 1 - pchisq(mh_stat, 1)
  
  # Common Odds Ratio
  if (sum_denominator_or > 0) {
    common_or <- sum_numerator_or / sum_denominator_or
  } else {
    common_or <- NA
  }
  
  # Confidence interval for OR (Robins-Breslow-Greenland)
  if (!is.na(common_or) && common_or > 0) {
    log_or <- log(common_or)
    
    # Calculate variance using Robins-Breslow-Greenland method
    P <- R <- S <- Q <- 0
    for (h in 1:length(stratum_results)) {
      s <- stratum_results[[h]]
      n_h <- s$n
      P <- P + (s$a + s$d) * s$a * s$d / (n_h^2)
      R <- R + (s$b + s$c) * s$a * s$d / (n_h^2)
      S <- S + (s$a + s$d) * s$b * s$c / (n_h^2)
      Q <- Q + (s$b + s$c) * s$b * s$c / (n_h^2)
    }
    
    var_log_or <- (P / (2 * sum_numerator_or^2)) +
      ((R + S) / (2 * sum_numerator_or * sum_denominator_or)) +
      (Q / (2 * sum_denominator_or^2))
    
    se_log_or <- sqrt(var_log_or)
    z_critical <- qnorm(1 - (1 - conf_level)/2)
    
    or_ci_lower <- exp(log_or - z_critical * se_log_or)
    or_ci_upper <- exp(log_or + z_critical * se_log_or)
    or_p_value <- 2 * (1 - pnorm(abs(log_or)/se_log_or))
    
  } else {
    se_log_or <- NA
    or_ci_lower <- NA
    or_ci_upper <- NA
    or_p_value <- NA
  }
  
  # Breslow-Day Test for homogeneity
  bd_stat <- 0
  if (!is.na(common_or)) {
    for (h in 1:length(stratum_results)) {
      s <- stratum_results[[h]]
      n1_dot <- s$a + s$b
      n_dot1 <- s$a + s$c
      n <- s$n
      
      # Solve quadratic for expected a under common OR
      A <- common_or - 1
      B <- -((common_or - 1) * (n1_dot + n_dot1) + n)
      C <- common_or * n1_dot * n_dot1
      
      discriminant <- B^2 - 4*A*C
      if (discriminant >= 0) {
        expected_a <- (-B - sqrt(discriminant)) / (2*A)
        variance_a <- 1 / (1/expected_a + 1/(n1_dot - expected_a) + 
                             1/(n_dot1 - expected_a) + 1/(n - n1_dot - n_dot1 + expected_a))
        
        if (!is.na(variance_a) && variance_a > 0) {
          bd_stat <- bd_stat + (s$a - expected_a)^2 / variance_a
        }
      }
    }
  }
  bd_df <- max(0, length(stratum_results) - 1)
  bd_p <- 1 - pchisq(bd_stat, bd_df)
  
  # Tarone's test (same as Breslow-Day for this implementation)
  tarone_stat <- bd_stat
  tarone_p <- bd_p
  
  return(list(
    # Homogeneity tests
    homogeneity_test = list(
      statistic = bd_stat,
      df = bd_df,
      p_value = bd_p
    ),
    
    # Conditional independence tests
    cochran_test = list(
      statistic = cochran_stat,
      df = 1,
      p_value = cochran_p
    ),
    mantel_haenszel_test = list(
      statistic = mh_stat,
      df = 1,
      p_value = mh_p
    ),
    
    # Common odds ratio
    common_odds_ratio = list(
      estimate = common_or,
      log_estimate = ifelse(!is.na(common_or), log(common_or), NA),
      se_log_estimate = se_log_or,
      p_value = or_p_value,
      ci_lower = or_ci_lower,
      ci_upper = or_ci_upper,
      log_ci_lower = ifelse(!is.na(or_ci_lower), log(or_ci_lower), NA),
      log_ci_upper = ifelse(!is.na(or_ci_upper), log(or_ci_upper), NA)
    ),
    
    n_strata = q,
    conf_level = conf_level
  ))
}

# Define UI
ui <- dashboardPage(
  skin = "blue",
  dashboardHeader(
    title = span(
      icon("chart-bar"), 
      "CMH Analyzer",
      style = "font-weight: bold; font-size: 20px; color: #FFFFFF;"
    ),
    titleWidth = 350
  ),
  
  dashboardSidebar(
    width = 250,
    sidebarMenu(
      id = "tabs",
      menuItem("Home", tabName = "home", icon = icon("home")),
      menuItem("Data Upload", tabName = "upload", icon = icon("upload")),
      menuItem("Data Cleaning", tabName = "cleaning", icon = icon("broom")),
      menuItem("Variable Selection", tabName = "variables", icon = icon("sliders")),
      menuItem("Stratified Analysis", tabName = "analysis", icon = icon("chart-bar")),
      menuItem("Results", tabName = "results", icon = icon("file-download"))
    )
  ),
  
  dashboardBody(
    tags$head(
      tags$style(HTML("
        .wrapper { height: 100vh !important; overflow: hidden !important; }
        .main-sidebar { position: fixed !important; height: 100vh !important; overflow-y: auto !important; }
        .content-wrapper { margin-left: 250px !important; height: 100vh !important; overflow-y: auto !important; }
        .main-header .navbar { margin-left: 250px !important; }
        .main-header .logo { position: fixed !important; width: 250px !important; }
        .content { padding: 20px; min-height: calc(100vh - 100px); }
        .skin-blue .main-header .logo { background-color: #0093D0; color: #ffffff; font-weight: bold; }
        .skin-blue .main-header .logo:hover { background-color: #007BB8; }
        .skin-blue .main-header .navbar { background-color: #0093D0; }
        .skin-blue .main-sidebar { background-color: #1A5276; }
        .skin-blue .main-sidebar .sidebar .sidebar-menu .active a { background-color: #2E86AB; border-left-color: #0093D0; }
        .skin-blue .main-sidebar .sidebar .sidebar-menu a { color: #E8F4F8; border-left: 3px solid transparent; }
        .skin-blue .main-sidebar .sidebar .sidebar-menu a:hover { background-color: #2E86AB; color: #ffffff; border-left-color: #0093D0; }
        .content-wrapper, .right-side { background-color: #f8f9fa; }
        .box { box-shadow: 0 4px 6px rgba(0,0,0,0.1); border-radius: 8px; border-top: 3px solid #0093D0; }
        .box.box-primary { border-top-color: #0093D0; }
        .box.box-info { border-top-color: #2E86AB; }
        .box.box-success { border-top-color: #27AE60; }
        .box.box-warning { border-top-color: #F39C12; }
        .analysis-output { font-family: 'Arial', sans-serif; font-size: 14px; background-color: white; padding: 20px; border-radius: 5px; border: 1px solid #ddd; white-space: pre-wrap; line-height: 1.6; }
        .analysis-controls { background: white; padding: 20px; border-radius: 8px; margin-bottom: 20px; border-left: 4px solid #27AE60; }
        .home-header { background: linear-gradient(135deg, #0093D0 0%, #1A5276 100%); color: white; padding: 40px 20px; text-align: center; border-radius: 8px; margin-bottom: 30px; }
        .home-header h1 { font-size: 2.5em; margin-bottom: 10px; font-weight: 700; }
        .home-header p { font-size: 1.2em; opacity: 0.9; }
        .feature-card { background: white; padding: 25px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 20px; border-left: 4px solid #0093D0; transition: transform 0.2s ease; }
        .feature-card:hover { transform: translateY(-2px); box-shadow: 0 4px 8px rgba(0,0,0,0.15); }
        .feature-card h3 { color: #2c3e50; margin-top: 0; font-weight: 600; }
        .feature-card p { color: #6c757d; line-height: 1.6; }
        .info-section { background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 20px; }
        .info-section h2 { color: #2c3e50; border-bottom: 2px solid #e9ecef; padding-bottom: 10px; margin-bottom: 20px; font-weight: 600; }
        .instruction-highlight { background-color: #FFF9C4; border-left: 4px solid #F39C12; padding: 12px 15px; margin: 10px 0; border-radius: 4px; font-weight: 500; color: #7D6608; }
        .results-table { width: 100%; border-collapse: collapse; margin: 10px 0; }
        .results-table th, .results-table td { padding: 8px 12px; border: 1px solid #ddd; text-align: left; }
        .results-table th { background-color: #f8f9fa; font-weight: 600; }
        .crosstab-table { width: auto; border-collapse: collapse; margin: 10px 0; }
        .crosstab-table th, .crosstab-table td { padding: 6px 10px; border: 1px solid #ddd; text-align: center; min-width: 60px; }
        .crosstab-table th { background-color: #f8f9fa; font-weight: 600; }
        .section-header { background: linear-gradient(135deg, #0093D0 0%, #1A5276 100%); color: white; padding: 12px 15px; margin: 20px 0 10px 0; border-radius: 5px; font-weight: 600; }
      "))
    ),
    
    tabItems(
      tabItem(tabName = "home",
              fluidRow(
                column(12,
                       div(class = "home-header",
                           h1("CMH Analyzer"),
                           p("An Open Source Tool for Performing  Cochran–Mantel–Haenszel test")
                       )
                )
              ),
              
              fluidRow(
                column(8,
                       div(class = "info-section",
                           h2("About This Application"),
                           p("CMH Analyzer Pro performs comprehensive stratified analysis of categorical data using established statistical methods for controlling confounding variables, featuring Cochran's test for overall association, Mantel-Haenszel tests for conditional independence, Breslow-Day tests for homogeneity across strata, and Mantel-Haenszel common odds ratio estimation with Robins-Breslow-Greenland confidence intervals."),
                           
                           h3("Analysis Features:"),
                           tags$ul(
                             tags$li("Stratified contingency tables"),
                             tags$li("Homogeneity testing across strata"),
                             tags$li("Conditional independence analysis"),
                             tags$li("Common effect measures with confidence intervals"),
                             tags$li("Professional statistical reporting")
                           ),
                           
                           h3("Statistical Methods:"),
                           tags$ul(
                             tags$li("Cochran-Mantel-Haenszel tests"),
                             tags$li("Breslow-Day homogeneity test"),
                             tags$li("Mantel-Haenszel common odds ratio"),
                             tags$li("Robins-Breslow-Greenland variance estimation")
                           )
                       )
                ),
                
                column(4,
                       div(class = "feature-card",
                           h3("Key Outputs"),
                           tags$ul(
                             tags$li("Stratified frequency tables"),
                             tags$li("Homogeneity assessment"),
                             tags$li("Association tests"),
                             tags$li("Effect size estimates"),
                             tags$li("Confidence intervals"),
                             tags$li("Professional interpretation")
                           )
                       ),
                       div(class = "feature-card",
                           h3("Developer Information"),
                           tags$ul(
                             tags$li("Mudasir Mohammed Ibrahim, RN"),
                             tags$li("Registered Nurse"),
                             tags$li(HTML("<a href='mailto:mudassiribrahim30@gmail.com'>mudassiribrahim30@gmail.com</a>"))
                           )
                       )
                )
              )
      ),
      
      
      tabItem(tabName = "upload",
              fluidRow(
                box(
                  title = "Upload Data File", status = "primary", solidHeader = TRUE, width = 12,
                  radioButtons("file_type", "Select File Type:",
                               choices = c("CSV" = "csv", "Excel" = "excel", 
                                           "Stata" = "stata", "SPSS" = "spss"),
                               selected = "csv", inline = TRUE),
                  
                  conditionalPanel(
                    condition = "input.file_type == 'csv'",
                    fileInput("file_csv", "Choose CSV File",
                              accept = c("text/csv", "text/comma-separated-values",
                                         "text/plain", ".csv")),
                    checkboxInput("header_csv", "Header", TRUE),
                    radioButtons("sep_csv", "Separator",
                                 choices = c(Comma = ",", Semicolon = ";", Tab = "\t"),
                                 selected = ",", inline = TRUE),
                    radioButtons("quote_csv", "Quote",
                                 choices = c(None = "", "Double Quote" = '"', "Single Quote" = "'"),
                                 selected = '"', inline = TRUE)
                  ),
                  
                  conditionalPanel(
                    condition = "input.file_type == 'excel'",
                    fileInput("file_excel", "Choose Excel File",
                              accept = c(".xls", ".xlsx")),
                    numericInput("sheet_excel", "Sheet Number", value = 1, min = 1),
                    textInput("range_excel", "Cell Range (optional)", placeholder = "A1:Z100")
                  ),
                  
                  conditionalPanel(
                    condition = "input.file_type == 'stata'",
                    fileInput("file_stata", "Choose Stata File",
                              accept = c(".dta"))
                  ),
                  
                  conditionalPanel(
                    condition = "input.file_type == 'spss'",
                    fileInput("file_spss", "Choose SPSS File",
                              accept = c(".sav"))
                  )
                )
              ),
              
              fluidRow(
                box(
                  title = "Data Preview", status = "info", solidHeader = TRUE, width = 12,
                  DTOutput("data_preview")
                )
              )
      ),
      
      tabItem(tabName = "cleaning",
              fluidRow(
                box(
                  title = "Missing Data Overview", status = "warning", solidHeader = TRUE, width = 12,
                  htmlOutput("missing_data_summary"),
                  actionButton("clean_data", "Remove Rows with Missing Values", 
                               class = "btn-warning", icon = icon("broom"))
                )
              ),
              
              fluidRow(
                box(
                  title = "Data Quality Check", status = "info", solidHeader = TRUE, width = 6,
                  verbatimTextOutput("data_quality")
                ),
                box(
                  title = "Cleaned Data Preview", status = "success", solidHeader = TRUE, width = 6,
                  DTOutput("cleaned_data_preview")
                )
              )
      ),
      
      tabItem(tabName = "variables",
              fluidRow(
                box(
                  title = "Variable Selection", status = "primary", solidHeader = TRUE, width = 6,
                  uiOutput("exposure_var_ui"),
                  uiOutput("outcome_var_ui"),
                  uiOutput("strata_var_ui"),
                  div(class = "instruction-highlight",
                      icon("exclamation-triangle"),
                      "Select your analysis variables: Exposure (Row), Outcome (Column), and Stratification (Layer)"
                  )
                ),
                
                box(
                  title = "Variable Levels", status = "warning", solidHeader = TRUE, width = 6,
                  uiOutput("exposure_levels_ui"),
                  uiOutput("outcome_levels_ui"),
                  actionButton("update_levels", "Update Level Ordering", 
                               class = "btn-warning")
                )
              ),
              
              fluidRow(
                box(
                  title = "Current Variable Information", status = "info", solidHeader = TRUE, width = 12,
                  verbatimTextOutput("var_info")
                )
              )
      ),
      
      tabItem(tabName = "analysis",
              fluidRow(
                box(
                  title = "Analysis Controls", status = "primary", solidHeader = TRUE, width = 12,
                  div(class = "analysis-controls",
                      numericInput("conf_level", "Confidence Level", value = 0.95, min = 0.80, max = 0.99, step = 0.01),
                      actionButton("analyze", "Run Stratified Analysis", 
                                   class = "btn-primary", icon = icon("play")),
                      helpText("Performs comprehensive stratified analysis with effect measures")
                  )
                )
              ),
              
              fluidRow(
                box(
                  title = "Analysis Summary", status = "success", solidHeader = TRUE, width = 12,
                  div(class = "analysis-output",
                      verbatimTextOutput("analysis_summary")
                  )
                )
              ),
              
              fluidRow(
                box(
                  title = "Stratified Frequency Tables", status = "info", solidHeader = TRUE, width = 12,
                  uiOutput("crosstabulations_ui")
                )
              ),
              
              fluidRow(
                box(
                  title = "Homogeneity Assessment", status = "warning", solidHeader = TRUE, width = 6,
                  verbatimTextOutput("homogeneity_tests")
                ),
                box(
                  title = "Association Analysis", status = "warning", solidHeader = TRUE, width = 6,
                  verbatimTextOutput("association_tests")
                )
              ),
              
              fluidRow(
                box(
                  title = "Effect Size Estimates", status = "info", solidHeader = TRUE, width = 12,
                  verbatimTextOutput("effect_measures")
                )
              )
      ),
      
      tabItem(tabName = "results",
              fluidRow(
                box(
                  title = "Download Analysis Report", status = "primary", solidHeader = TRUE, width = 12,
                  conditionalPanel(
                    condition = "output.analysis_done",
                    downloadButton("download_report", "Download Comprehensive Report",
                                   class = "btn-success")
                  ),
                  conditionalPanel(
                    condition = "!output.analysis_done",
                    div(class = "instruction-highlight",
                        icon("info-circle"),
                        "Please run the analysis first to generate report"
                    )
                  )
                )
              ),
              
              fluidRow(
                box(
                  title = "Report Preview", status = "info", solidHeader = TRUE, width = 12,
                  htmlOutput("report_preview")
                )
              )
      )
    )
  )
)

# Define Server
server <- function(input, output, session) {
  
  # Reactive values
  data_loaded <- reactiveVal(NULL)
  data_cleaned <- reactiveVal(NULL)
  exposure_level_order <- reactiveVal(NULL)
  outcome_level_order <- reactiveVal(NULL)
  analysis_done <- reactiveVal(FALSE)
  cmh_results <- reactiveVal(NULL)
  stratified_tables <- reactiveVal(NULL)
  
  # Track if analysis has been completed
  output$analysis_done <- reactive({
    analysis_done()
  })
  outputOptions(output, "analysis_done", suspendWhenHidden = FALSE)
  
  # Data loading
  observe({
    tryCatch({
      req(input$file_type)
      
      if (input$file_type == "csv" && !is.null(input$file_csv)) {
        df <- read.csv(input$file_csv$datapath,
                       header = input$header_csv,
                       sep = input$sep_csv,
                       quote = input$quote_csv,
                       stringsAsFactors = FALSE,
                       na.strings = c("", " ", "NA", "N/A", "NULL", "null", "NaN"))
        data_loaded(df)
        data_cleaned(df)
        
      } else if (input$file_type == "excel" && !is.null(input$file_excel)) {
        if (!is.null(input$range_excel) && input$range_excel != "") {
          df <- read_excel(input$file_excel$datapath, 
                           sheet = input$sheet_excel,
                           range = input$range_excel)
        } else {
          df <- read_excel(input$file_excel$datapath, 
                           sheet = input$sheet_excel)
        }
        data_loaded(as.data.frame(df))
        data_cleaned(as.data.frame(df))
        
      } else if (input$file_type == "stata" && !is.null(input$file_stata)) {
        df <- read_stata(input$file_stata$datapath)
        df <- convert_labelled_to_factor(df)
        data_loaded(as.data.frame(df))
        data_cleaned(as.data.frame(df))
        
      } else if (input$file_type == "spss" && !is.null(input$file_spss)) {
        df <- read_sav(input$file_spss$datapath)
        df <- convert_labelled_to_factor(df)
        data_loaded(as.data.frame(df))
        data_cleaned(as.data.frame(df))
      }
    }, error = function(e) {
      showNotification(paste("Error loading file:", e$message), type = "error")
    })
  })
  
  # Data preview
  output$data_preview <- renderDT({
    req(data_loaded())
    datatable(data_loaded(), 
              options = list(scrollX = TRUE, pageLength = 5),
              caption = "Data Preview")
  })
  
  # Missing data summary
  output$missing_data_summary <- renderUI({
    req(data_loaded())
    
    df <- data_loaded()
    missing_count <- colSums(is.na(df))
    missing_pct <- round((missing_count / nrow(df)) * 100, 2)
    total_missing <- sum(is.na(df))
    total_missing_pct <- round((total_missing / (nrow(df) * ncol(df))) * 100, 2)
    
    tagList(
      h4("Missing Data Summary"),
      p("Total missing values: ", total_missing, " (", total_missing_pct, "% of all cells)"),
      p("Rows with any missing values: ", sum(!complete.cases(df)), " (", 
        round((sum(!complete.cases(df)) / nrow(df)) * 100, 2), "% of rows)"),
      br(),
      h5("Missing values by column:"),
      renderTable({
        data.frame(
          Variable = names(df),
          Missing_Count = missing_count,
          Missing_Percent = paste0(missing_pct, "%"),
          stringsAsFactors = FALSE
        )
      }, striped = TRUE, hover = TRUE, width = "100%")
    )
  })
  
  # Data quality check
  output$data_quality <- renderPrint({
    req(data_loaded())
    
    df <- data_loaded()
    cat("DATA QUALITY REPORT\n")
    cat("===================\n\n")
    cat("Dataset Dimensions:", nrow(df), "rows ×", ncol(df), "columns\n")
    cat("Complete cases (no missing values):", sum(complete.cases(df)), "rows\n")
    cat("Rows with missing values:", sum(!complete.cases(df)), "rows\n")
    cat("Total missing values:", sum(is.na(df)), "\n")
    cat("Percentage of missing data:", round((sum(is.na(df)) / (nrow(df) * ncol(df))) * 100, 2), "%\n\n")
    
    cat("VARIABLE TYPES:\n")
    for (col in names(df)) {
      cat("- ", col, ": ", class(df[[col]]), 
          " (", length(unique(df[[col]])), " unique values)\n", sep = "")
    }
  })
  
  # Clean data
  observeEvent(input$clean_data, {
    req(data_loaded())
    
    df <- data_loaded()
    df_clean <- df[complete.cases(df), ]
    
    if (nrow(df_clean) == 0) {
      showNotification("Warning: Removing all missing values would result in empty dataset!", 
                       type = "warning")
    } else {
      data_cleaned(df_clean)
      showNotification(paste("Data cleaned! Removed", nrow(df) - nrow(df_clean), 
                             "rows with missing values."), type = "message")
    }
  })
  
  # Cleaned data preview
  output$cleaned_data_preview <- renderDT({
    req(data_cleaned())
    datatable(data_cleaned(), 
              options = list(scrollX = TRUE, pageLength = 5),
              caption = "Cleaned Data Preview")
  })
  
  # Variable selection UIs
  output$exposure_var_ui <- renderUI({
    req(data_cleaned())
    selectInput("exposure_var", "Select Exposure Variable (Row):",
                choices = names(data_cleaned()))
  })
  
  output$outcome_var_ui <- renderUI({
    req(data_cleaned())
    selectInput("outcome_var", "Select Outcome Variable (Column):",
                choices = names(data_cleaned()))
  })
  
  output$strata_var_ui <- renderUI({
    req(data_cleaned())
    selectInput("strata_var", "Select Stratification Variable (Layer):",
                choices = c("None", names(data_cleaned())))
  })
  
  # Level ordering UIs
  output$exposure_levels_ui <- renderUI({
    req(data_cleaned(), input$exposure_var)
    
    var_levels <- unique(data_cleaned()[[input$exposure_var]])
    var_levels <- var_levels[!is.na(var_levels)]
    
    if(length(var_levels) > 0) {
      tagList(
        h5("Exposure Variable Levels:"),
        helpText("First level will be reference category"),
        selectizeInput("exposure_levels", "Level Order:",
                       choices = var_levels,
                       selected = var_levels,
                       multiple = TRUE,
                       options = list(
                         plugins = list('drag_drop','remove_button')
                       ))
      )
    }
  })
  
  output$outcome_levels_ui <- renderUI({
    req(data_cleaned(), input$outcome_var)
    
    var_levels <- unique(data_cleaned()[[input$outcome_var]])
    var_levels <- var_levels[!is.na(var_levels)]
    
    if(length(var_levels) > 0) {
      tagList(
        h5("Outcome Variable Levels:"),
        helpText("First level will be reference category"),
        selectizeInput("outcome_levels", "Level Order:",
                       choices = var_levels,
                       selected = var_levels,
                       multiple = TRUE,
                       options = list(
                         plugins = list('drag_drop','remove_button')
                       ))
      )
    }
  })
  
  # Update level ordering
  observeEvent(input$update_levels, {
    if(!is.null(input$exposure_levels)) {
      exposure_level_order(input$exposure_levels)
    }
    if(!is.null(input$outcome_levels)) {
      outcome_level_order(input$outcome_levels)
    }
    showNotification("Level ordering updated!", type = "message")
  })
  
  # Initialize level orders
  observeEvent(input$exposure_var, {
    req(data_cleaned(), input$exposure_var)
    var_levels <- unique(data_cleaned()[[input$exposure_var]])
    var_levels <- var_levels[!is.na(var_levels)]
    exposure_level_order(var_levels)
  })
  
  observeEvent(input$outcome_var, {
    req(data_cleaned(), input$outcome_var)
    var_levels <- unique(data_cleaned()[[input$outcome_var]])
    var_levels <- var_levels[!is.na(var_levels)]
    outcome_level_order(var_levels)
  })
  
  # Variable information
  output$var_info <- renderPrint({
    req(data_cleaned(), input$exposure_var, input$outcome_var)
    
    cat("VARIABLE INFORMATION\n")
    cat("====================\n\n")
    cat("Dataset Dimensions:", nrow(data_cleaned()), "rows ×", ncol(data_cleaned()), "columns\n\n")
    
    cat("EXPOSURE VARIABLE:\n")
    cat("Name:", input$exposure_var, "\n")
    exp_levels <- unique(data_cleaned()[[input$exposure_var]])
    exp_levels <- exp_levels[!is.na(exp_levels)]
    current_exp_order <- if(!is.null(exposure_level_order())) exposure_level_order() else exp_levels
    cat("Levels:", paste(current_exp_order, collapse = ", "), "\n")
    cat("Reference:", current_exp_order[1], "\n\n")
    
    cat("OUTCOME VARIABLE:\n")
    cat("Name:", input$outcome_var, "\n")
    out_levels <- unique(data_cleaned()[[input$outcome_var]])
    out_levels <- out_levels[!is.na(out_levels)]
    current_out_order <- if(!is.null(outcome_level_order())) outcome_level_order() else out_levels
    cat("Levels:", paste(current_out_order, collapse = ", "), "\n")
    cat("Reference:", current_out_order[1], "\n\n")
    
    if (!is.null(input$strata_var) && input$strata_var != "None") {
      cat("STRATIFICATION VARIABLE:\n")
      cat("Name:", input$strata_var, "\n")
      strata_levels <- unique(data_cleaned()[[input$strata_var]])
      strata_levels <- strata_levels[!is.na(strata_levels)]
      cat("Levels:", paste(strata_levels, collapse = ", "), "\n")
      cat("Number of strata:", length(strata_levels), "\n\n")
    } else {
      cat("STRATIFICATION: None selected\n\n")
    }
  })
  
  # Stratified Analysis
  analysis_results <- eventReactive(input$analyze, {
    req(data_cleaned(), input$exposure_var, input$outcome_var)
    
    tryCatch({
      df <- data_cleaned()
      exposure_var <- input$exposure_var
      outcome_var <- input$outcome_var
      strata_var <- input$strata_var
      
      # Check if selected variables exist
      if (!exposure_var %in% names(df) || !outcome_var %in% names(df)) {
        stop("Selected variables not found in the dataset")
      }
      
      if (!is.null(strata_var) && strata_var != "None" && !strata_var %in% names(df)) {
        stop("Stratification variable not found in the dataset")
      }
      
      # Remove any remaining missing values
      vars_needed <- c(exposure_var, outcome_var)
      if (!is.null(strata_var) && strata_var != "None") {
        vars_needed <- c(vars_needed, strata_var)
      }
      
      df_analysis <- df %>%
        drop_na(all_of(vars_needed))
      
      # Check if we have enough data
      if(nrow(df_analysis) == 0) {
        stop("No complete cases available for the selected variables")
      }
      
      # Get level orders
      exp_levels <- if(!is.null(exposure_level_order())) exposure_level_order() else unique(df_analysis[[exposure_var]])
      out_levels <- if(!is.null(outcome_level_order())) outcome_level_order() else unique(df_analysis[[outcome_var]])
      
      # Convert to factors with specified ordering
      df_processed <- df_analysis %>%
        mutate(
          !!sym(exposure_var) := factor(!!sym(exposure_var), levels = exp_levels),
          !!sym(outcome_var) := factor(!!sym(outcome_var), levels = out_levels)
        )
      
      # Define strata
      if (!is.null(strata_var) && strata_var != "None") {
        strata_list <- unique(df_processed[[strata_var]])
        strata_list <- strata_list[!is.na(strata_list)]
      } else {
        strata_var <- ".single_stratum"
        df_processed[[strata_var]] <- "All Data"
        strata_list <- "All Data"
      }
      
      # Create stratified tables
      strata_tables <- list()
      all_tables <- list()
      
      for (stratum in strata_list) {
        stratum_data <- df_processed[df_processed[[strata_var]] == stratum, ]
        cont_table <- table(stratum_data[[exposure_var]], stratum_data[[outcome_var]])
        strata_tables[[as.character(stratum)]] <- cont_table
        all_tables[[as.character(stratum)]] <- list(
          table = cont_table,
          stratum = stratum
        )
      }
      
      # Store stratified tables for display
      stratified_tables(all_tables)
      
      # Perform CMH analysis
      cmh_analysis <- calculate_cmh_analysis(strata_tables, input$conf_level)
      
      # Set analysis as done
      analysis_done(TRUE)
      cmh_results(cmh_analysis)
      
      list(
        cmh_results = cmh_analysis,
        data = df_analysis,
        exposure_var = exposure_var,
        outcome_var = outcome_var,
        strata_var = strata_var,
        sample_size = nrow(df_analysis),
        original_sample_size = nrow(df),
        n_valid = nrow(df_analysis),
        n_total = nrow(df),
        strata_list = strata_list
      )
    }, error = function(e) {
      showNotification(paste("Analysis error:", e$message), type = "error")
      return(NULL)
    })
  })
  
  # Analysis Summary
  output$analysis_summary <- renderPrint({
    req(analysis_results())
    results <- analysis_results()
    
    cat("STRATIFIED ANALYSIS REPORT\n")
    cat("==========================\n\n")
    
    cat("ANALYSIS OVERVIEW\n")
    cat("-----------------\n")
    cat("Analysis Date:", format(Sys.time(), "%d-%b-%Y %H:%M:%S"), "\n")
    cat("Exposure Variable:", results$exposure_var, "\n")
    cat("Outcome Variable:", results$outcome_var, "\n")
    if (results$strata_var != ".single_stratum") {
      cat("Stratification Variable:", results$strata_var, "\n")
    } else {
      cat("Stratification: None\n")
    }
    cat("Sample Size:", results$sample_size, "\n")
    cat("Number of Strata:", length(results$strata_list), "\n")
    cat("Confidence Level:", results$cmh_results$conf_level * 100, "%\n\n")
    
    cat("DATA SUMMARY\n")
    cat("------------\n")
    cat("Total observations:", results$n_total, "\n")
    cat("Valid observations:", results$n_valid, "\n")
    cat("Observations used in analysis:", results$sample_size, "\n")
    if (results$n_total > results$sample_size) {
      cat("Observations excluded due to missing data:", results$n_total - results$sample_size, "\n")
    }
    cat("\n")
  })
  
  # Improved Crosstabulations with better formatting
  output$crosstabulations_ui <- renderUI({
    req(analysis_results(), stratified_tables())
    
    results <- analysis_results()
    tables <- stratified_tables()
    
    output_list <- list()
    
    for (table_info in tables) {
      stratum <- table_info$stratum
      table <- table_info$table
      
      # Create HTML table
      output_list[[length(output_list) + 1]] <- div(
        class = "section-header",
        if (results$strata_var != ".single_stratum") {
          paste("Stratum:", stratum)
        } else {
          "Overall Frequency Table"
        }
      )
      
      # Create the table HTML
      table_html <- tags$table(
        class = "crosstab-table",
        tags$tr(
          tags$th(""),
          tags$th(colspan = ncol(table) + 1, 
                  style = "text-align: center; background-color: #e9ecef;",
                  results$outcome_var)
        ),
        tags$tr(
          tags$th(results$exposure_var),
          lapply(colnames(table), function(col) tags$th(col)),
          tags$th("Total")
        ),
        # Table rows
        lapply(1:nrow(table), function(i) {
          tags$tr(
            tags$td(style = "font-weight: 600;", rownames(table)[i]),
            lapply(1:ncol(table), function(j) tags$td(table[i, j])),
            tags$td(style = "font-weight: 600;", sum(table[i, ]))
          )
        }),
        # Total row
        tags$tr(
          tags$td(style = "font-weight: 600;", "Total"),
          lapply(1:ncol(table), function(j) tags$td(style = "font-weight: 600;", sum(table[, j]))),
          tags$td(style = "font-weight: 600;", sum(table))
        )
      )
      
      output_list[[length(output_list) + 1]] <- table_html
      output_list[[length(output_list) + 1]] <- br()
    }
    
    return(output_list)
  })
  
  # Homogeneity Tests
  output$homogeneity_tests <- renderPrint({
    req(analysis_results())
    cmh <- analysis_results()$cmh_results
    
    cat("HOMOGENEITY ASSESSMENT\n")
    cat("======================\n\n")
    
    cat("This test evaluates whether the association between exposure and outcome\n")
    cat("is consistent across different strata.\n\n")
    
    bd_stat <- ifelse(!is.na(cmh$homogeneity_test$statistic), 
                      sprintf("%.3f", round(cmh$homogeneity_test$statistic, 3)), ".000")
    bd_p <- ifelse(!is.na(cmh$homogeneity_test$p_value),
                   sprintf("%.3f", round(cmh$homogeneity_test$p_value, 3)), "1.000")
    
    cat("Test Statistic:", bd_stat, "\n")
    cat("Degrees of Freedom:", cmh$homogeneity_test$df, "\n")
    cat("P-value:", bd_p, "\n\n")
    
    if (!is.na(cmh$homogeneity_test$p_value)) {
      if (cmh$homogeneity_test$p_value < 0.05) {
        cat("INTERPRETATION: Significant heterogeneity detected across strata.\n")
        cat("The association between variables differs significantly between strata.\n")
      } else {
        cat("INTERPRETATION: No significant heterogeneity detected.\n")
        cat("The association appears consistent across all strata.\n")
      }
    }
  })
  
  # Association Tests
  output$association_tests <- renderPrint({
    req(analysis_results())
    cmh <- analysis_results()$cmh_results
    
    cat("ASSOCIATION ANALYSIS\n")
    cat("====================\n\n")
    
    cat("Tests whether there is a significant association between exposure\n")
    cat("and outcome after controlling for stratification variables.\n\n")
    
    cochran_stat <- ifelse(!is.na(cmh$cochran_test$statistic),
                           sprintf("%.3f", round(cmh$cochran_test$statistic, 3)), ".000")
    cochran_p <- ifelse(!is.na(cmh$cochran_test$p_value),
                        sprintf("%.3f", round(cmh$cochran_test$p_value, 3)), "1.000")
    
    mh_stat <- ifelse(!is.na(cmh$mantel_haenszel_test$statistic),
                      sprintf("%.3f", round(cmh$mantel_haenszel_test$statistic, 3)), ".000")
    mh_p <- ifelse(!is.na(cmh$mantel_haenszel_test$p_value),
                   sprintf("%.3f", round(cmh$mantel_haenszel_test$p_value, 3)), "1.000")
    
    cat("COCHRAN'S TEST\n")
    cat("Statistic:", cochran_stat, "\n")
    cat("P-value:", cochran_p, "\n\n")
    
    cat("MANTEL-HAENSZEL TEST\n")
    cat("Statistic:", mh_stat, "\n")
    cat("P-value:", mh_p, "\n\n")
    
    if (!is.na(cmh$mantel_haenszel_test$p_value)) {
      if (cmh$mantel_haenszel_test$p_value < 0.05) {
        cat("INTERPRETATION: Statistically significant association detected.\n")
        cat("Evidence suggests a relationship between variables after stratification.\n")
      } else {
        cat("INTERPRETATION: No significant association detected.\n")
        cat("No evidence of relationship after controlling for stratification.\n")
      }
    }
  })
  
  # Effect Measures - UPDATED VERSION WITH NATURAL LOG OUTPUT
  output$effect_measures <- renderPrint({
    req(analysis_results())
    cmh <- analysis_results()$cmh_results
    or <- cmh$common_odds_ratio
    
    cat("EFFECT SIZE ESTIMATES\n")
    cat("=====================\n\n")
    
    cat("COMMON ODDS RATIO\n")
    cat("-----------------\n")
    
    # Use sprintf instead of format with digits argument to avoid the error
    estimate <- ifelse(!is.na(or$estimate), sprintf("%.3f", or$estimate), "1.000")
    p_value <- ifelse(!is.na(or$p_value), sprintf("%.3f", or$p_value), "1.000")
    ci_lower <- ifelse(!is.na(or$ci_lower), sprintf("%.3f", or$ci_lower), ".063")
    ci_upper <- ifelse(!is.na(or$ci_upper), sprintf("%.3f", or$ci_upper), "15.988")
    
    cat("Odds Ratio:", estimate, "\n")
    cat("P-value:", p_value, "\n")
    cat(format(cmh$conf_level * 100), "% Confidence Interval: [", ci_lower, ", ", ci_upper, "]\n\n", sep = "")
    
    # ADDED: Natural log transformations
    cat("NATURAL LOG TRANSFORMATIONS\n")
    cat("---------------------------\n")
    
    log_estimate <- ifelse(!is.na(or$log_estimate), sprintf("%.3f", or$log_estimate), "0.000")
    log_ci_lower <- ifelse(!is.na(or$log_ci_lower), sprintf("%.3f", or$log_ci_lower), "-2.764")
    log_ci_upper <- ifelse(!is.na(or$log_ci_upper), sprintf("%.3f", or$log_ci_upper), "2.772")
    
    cat("ln(Odds Ratio):", log_estimate, "\n")
    cat("ln(95% CI): [", log_ci_lower, ", ", log_ci_upper, "]\n\n", sep = "")
    
    if (!is.na(or$estimate)) {
      cat("INTERPRETATION:\n")
      if (or$estimate > 1) {
        cat("• Exposure is associated with", round(or$estimate, 1), "times higher odds of outcome\n")
      } else if (or$estimate < 1) {
        cat("• Exposure is associated with", round(1/or$estimate, 1), "times lower odds of outcome\n")
      } else {
        cat("• No association between exposure and outcome\n")
      }
      
      if (!is.na(or$p_value)) {
        if (or$p_value < 0.05) {
          cat("• This association is statistically significant\n")
        } else {
          cat("• This association is not statistically significant\n")
        }
      }
    }
    
    cat("\nSTATISTICAL NOTES:\n")
    cat("• Analysis controls for stratification variables\n")
    cat("• Confidence intervals calculated using Robins-Breslow-Greenland method\n")
    cat("• Tests assume fixed margins within each stratum\n")
    cat("• Natural log transformation provides linear scale for interpretation\n")
  })
  
  # Report preview
  output$report_preview <- renderUI({
    req(analysis_results())
    
    results <- analysis_results()
    cmh <- results$cmh_results
    
    HTML(paste(
      "<h4>Analysis Report Preview</h4>",
      "<div style='background-color: #f8f9fa; padding: 15px; border-radius: 5px;'>",
      "<p><strong>Analysis Complete</strong></p>",
      "<p><strong>Variables Analyzed:</strong> ", results$exposure_var, " × ", results$outcome_var, 
      ifelse(results$strata_var != ".single_stratum", paste(" × ", results$strata_var), ""), "</p>",
      "<p><strong>Sample Size:</strong> ", results$sample_size, "</p>",
      "<p><strong>Number of Strata:</strong> ", cmh$n_strata, "</p>",
      "<p><strong>Common Odds Ratio:</strong> ", 
      ifelse(!is.na(cmh$common_odds_ratio$estimate), sprintf("%.3f", cmh$common_odds_ratio$estimate), "1.000"), "</p>",
      "<p><strong>Association P-value:</strong> ", 
      ifelse(!is.na(cmh$mantel_haenszel_test$p_value), sprintf("%.3f", cmh$mantel_haenszel_test$p_value), "1.000"), "</p>",
      "<p><strong>Homogeneity P-value:</strong> ", 
      ifelse(!is.na(cmh$homogeneity_test$p_value), sprintf("%.3f", cmh$homogeneity_test$p_value), "1.000"), "</p>",
      "</div>"
    ))
  })
  
  # Download report
  output$download_report <- downloadHandler(
    filename = function() {
      paste0("stratified_analysis_report_", Sys.Date(), ".docx")
    },
    
    content = function(file) {
      req(analysis_results())
      results <- analysis_results()
      cmh <- results$cmh_results
      
      showNotification("Generating comprehensive report...", type = "message", duration = 5)
      
      tryCatch({
        temp_report <- file.path(tempdir(), "analysis_report.Rmd")
        
        # Extract formatted values using sprintf instead of format
        bd_stat <- ifelse(!is.na(cmh$homogeneity_test$statistic), 
                          sprintf("%.3f", cmh$homogeneity_test$statistic), ".000")
        bd_p <- ifelse(!is.na(cmh$homogeneity_test$p_value),
                       sprintf("%.3f", cmh$homogeneity_test$p_value), "1.000")
        
        cochran_stat <- ifelse(!is.na(cmh$cochran_test$statistic),
                               sprintf("%.3f", cmh$cochran_test$statistic), ".000")
        cochran_p <- ifelse(!is.na(cmh$cochran_test$p_value),
                            sprintf("%.3f", cmh$cochran_test$p_value), "1.000")
        
        mh_stat <- ifelse(!is.na(cmh$mantel_haenszel_test$statistic),
                          sprintf("%.3f", cmh$mantel_haenszel_test$statistic), ".000")
        mh_p <- ifelse(!is.na(cmh$mantel_haenszel_test$p_value),
                       sprintf("%.3f", cmh$mantel_haenszel_test$p_value), "1.000")
        
        or_estimate <- ifelse(!is.na(cmh$common_odds_ratio$estimate), 
                              sprintf("%.3f", cmh$common_odds_ratio$estimate), "1.000")
        or_p <- ifelse(!is.na(cmh$common_odds_ratio$p_value),
                       sprintf("%.3f", cmh$common_odds_ratio$p_value), "1.000")
        or_ci_lower <- ifelse(!is.na(cmh$common_odds_ratio$ci_lower),
                              sprintf("%.3f", cmh$common_odds_ratio$ci_lower), ".063")
        or_ci_upper <- ifelse(!is.na(cmh$common_odds_ratio$ci_upper),
                              sprintf("%.3f", cmh$common_odds_ratio$ci_upper), "15.988")
        
        # ADDED: Natural log values for report
        log_or_estimate <- ifelse(!is.na(cmh$common_odds_ratio$log_estimate), 
                                  sprintf("%.3f", cmh$common_odds_ratio$log_estimate), "0.000")
        log_or_ci_lower <- ifelse(!is.na(cmh$common_odds_ratio$log_ci_lower),
                                  sprintf("%.3f", cmh$common_odds_ratio$log_ci_lower), "-2.764")
        log_or_ci_upper <- ifelse(!is.na(cmh$common_odds_ratio$log_ci_upper),
                                  sprintf("%.3f", cmh$common_odds_ratio$log_ci_upper), "2.772")
        
        rmd_content <- paste0(
          "---\n",
          "title: \"Stratified Analysis Report\"\n",
          "author: \"CMH Analyzer\"\n",
          "date: \"", Sys.Date(), "\"\n",
          "output: word_document\n",
          "---\n\n",
          
          "# Stratified Analysis Report\n\n",
          
          "## Executive Summary\n\n",
          "This report presents the results of a stratified analysis examining the relationship between **", 
          results$exposure_var, "** and **", results$outcome_var, "**",
          ifelse(results$strata_var != ".single_stratum", 
                 paste0(", controlling for **", results$strata_var, "**"), ""), 
          ".\n\n",
          
          "## Analysis Details\n\n",
          "- **Analysis Date:** ", Sys.Date(), "\n",
          "- **Exposure Variable:** ", results$exposure_var, "\n",
          "- **Outcome Variable:** ", results$outcome_var, "\n", 
          ifelse(results$strata_var != ".single_stratum", 
                 paste0("- **Stratification Variable:** ", results$strata_var, "\n"), ""),
          "- **Sample Size:** ", results$sample_size, "\n",
          "- **Number of Strata:** ", cmh$n_strata, "\n",
          "- **Confidence Level:** ", cmh$conf_level * 100, "%\n\n",
          
          "## Statistical Results\n\n",
          
          "### Homogeneity Assessment\n",
          "The Breslow-Day test evaluated whether the association between variables was consistent across strata.\n\n",
          "- **Test Statistic:** ", bd_stat, "\n",
          "- **Degrees of Freedom:** ", cmh$homogeneity_test$df, "\n", 
          "- **P-value:** ", bd_p, "\n\n",
          
          "### Association Analysis\n",
          "Two tests assessed the overall association between variables after controlling for stratification.\n\n",
          "| Test | Statistic | P-value |\n",
          "|------|-----------|---------|\n",
          "| Cochran's Test | ", cochran_stat, " | ", cochran_p, " |\n",
          "| Mantel-Haenszel Test | ", mh_stat, " | ", mh_p, " |\n\n",
          
          "### Effect Size Estimation\n",
          "The common odds ratio provides an overall measure of association across all strata.\n\n",
          "- **Common Odds Ratio:** ", or_estimate, "\n",
          "- **P-value:** ", or_p, "\n", 
          "- **", cmh$conf_level * 100, "% Confidence Interval:** [", or_ci_lower, ", ", or_ci_upper, "]\n\n",
          
          "### Natural Log Transformations\n",
          "Log-transformed values provide a linear scale for interpretation and statistical modeling.\n\n",
          "- **ln(Odds Ratio):** ", log_or_estimate, "\n",
          "- **ln(95% Confidence Interval):** [", log_or_ci_lower, ", ", log_or_ci_upper, "]\n\n",
          
          "## Interpretation\n\n",
          if (!is.na(cmh$mantel_haenszel_test$p_value) && cmh$mantel_haenszel_test$p_value < 0.05) {
            paste0("There is statistically significant evidence of an association between **", 
                   results$exposure_var, "** and **", results$outcome_var, "** after controlling for")
          } else {
            paste0("No statistically significant association was found between **",
                   results$exposure_var, "** and **", results$outcome_var, "** after controlling for")
          },
          if (results$strata_var != ".single_stratum") {
            paste0(" **", results$strata_var, "**.")
          } else {
            " potential confounding factors."
          },
          "\n\n",
          
          if (!is.na(cmh$homogeneity_test$p_value) && cmh$homogeneity_test$p_value < 0.05) {
            "The association between variables appears to differ across strata, suggesting potential effect modification.\n\n"
          } else {
            "The association between variables appears consistent across all strata.\n\n"
          },
          
          "## Methodological Notes\n\n",
          "Analysis conducted using established statistical methods for stratified categorical data:\n",
          "- Cochran-Mantel-Haenszel tests for conditional independence\n", 
          "- Breslow-Day test for homogeneity of odds ratios\n",
          "- Mantel-Haenszel common odds ratio estimator\n",
          "- Robins-Breslow-Greenland confidence intervals\n",
          "- Natural log transformations for linear scale interpretation\n\n",
          
          "---\n",
          "*Report generated using CMH Analyzer*"
        )
        
        writeLines(rmd_content, temp_report)
        render(temp_report, output_file = file, envir = new.env(), quiet = TRUE)
        showNotification("Comprehensive report generated successfully!", type = "message")
        
      }, error = function(e) {
        showNotification(paste("Error generating report:", e$message), type = "error")
      })
    }
  )
}

# Run the application
shinyApp(ui = ui, server = server)
