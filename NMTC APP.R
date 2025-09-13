# app.R
# SPSS-Style Statistical Analysis Software
library(shiny)
library(shinydashboard)
library(DT)
library(ggplot2)
library(plotly)
library(readxl)
library(haven)
library(dplyr)
library(tidyr)
library(car)
library(lmtest)
library(broom)
library(corrplot)
library(psych)
library(nortest)
library(shinycssloaders)
library(caret)
library(cluster)
library(FactoMineR)
library(factoextra)

# SPSS Color Scheme
spss_blue <- "#326EA6"
spss_dark_blue <- "#1A3A5F"
spss_light_gray <- "#F2F2F2"
spss_dark_gray <- "#7F7F7F"

# UI Definition - SPSS Style Interface
ui <- dashboardPage(
  skin = "blue",
  dashboardHeader(
    title = span(icon("chart-bar"), "SPSS Statistics", style = paste0("color: white; font-weight: bold; font-size: 18px; background-color:", spss_blue)),
    titleWidth = 250,
    dropdownMenu(
      type = "messages",
      badgeStatus = NULL,
      icon = icon("question-circle"),
      messageItem(
        from = "Help",
        message = "SPSS Statistics Help",
        icon = icon("question")
      )
    )
  ),
  
  dashboardSidebar(
    width = 250,
    sidebarMenu(
      id = "tabs",
      style = paste0("background-color:", spss_light_gray),
      menuItem("File", tabName = "file", icon = icon("folder-open"), 
               menuSubItem("Open", tabName = "open"),
               menuSubItem("Save", tabName = "save")),
      menuItem("Edit", tabName = "edit", icon = icon("edit")),
      menuItem("View", tabName = "view", icon = icon("desktop")),
      menuItem("Data", tabName = "data", icon = icon("database"),
               menuSubItem("Define Variable Properties", tabName = "define_var"),
               menuSubItem("Sort Cases", tabName = "sort_cases"),
               menuSubItem("Select Cases", tabName = "select_cases"),
               menuSubItem("Weight Cases", tabName = "weight_cases")),
      menuItem("Transform", tabName = "transform", icon = icon("exchange-alt"),
               menuSubItem("Compute Variable", tabName = "compute"),
               menuSubItem("Recode", tabName = "recode"),
               menuSubItem("Visual Binning", tabName = "binning")),
      menuItem("Analyze", tabName = "analyze", icon = icon("chart-bar"),
               menuItem("Descriptive Statistics", tabName = "descriptive", icon = icon("table"),
                        menuSubItem("Frequencies", tabName = "frequencies"),
                        menuSubItem("Descriptives", tabName = "descriptives"),
                        menuSubItem("Explore", tabName = "explore")
               ),
               menuItem("Compare Means", tabName = "compare_means", icon = icon("balance-scale"),
                        menuSubItem("Means", tabName = "means"),
                        menuSubItem("One-Sample T Test", tabName = "one_sample_ttest"),
                        menuSubItem("Independent Samples T Test", tabName = "independent_ttest"),
                        menuSubItem("Paired Samples T Test", tabName = "paired_ttest"),
                        menuSubItem("ANOVA", tabName = "anova_menu")
               ),
               menuItem("Correlation", tabName = "correlation_menu", icon = icon("random"),
                        menuSubItem("Bivariate", tabName = "bivariate"),
                        menuSubItem("Partial", tabName = "partial")
               ),
               menuItem("Regression", tabName = "regression_menu", icon = icon("chart-line"),
                        menuSubItem("Linear", tabName = "linear_reg"),
                        menuSubItem("Binary Logistic", tabName = "logistic_reg")
               ),
               menuItem("Dimension Reduction", tabName = "dim_reduction", icon = icon("compress"),
                        menuSubItem("Factor", tabName = "factor")
               ),
               menuItem("Scale", tabName = "scale_menu", icon = icon("sliders-h"),
                        menuSubItem("Reliability Analysis", tabName = "reliability")
               ),
               menuItem("Nonparametric Tests", tabName = "nonparametric", icon = icon("check-square"),
                        menuSubItem("Chi-Square", tabName = "chi_square"),
                        menuSubItem("2 Independent Samples", tabName = "two_independent_samples")
               )
      ),
      menuItem("Graphs", tabName = "graphs", icon = icon("chart-area"),
               menuSubItem("Chart Builder", tabName = "chart_builder"),
               menuSubItem("Graphboard Template Chooser", tabName = "graphboard")),
      menuItem("Utilities", tabName = "utilities", icon = icon("cog")),
      menuItem("Extensions", tabName = "extensions", icon = icon("puzzle-piece")),
      menuItem("Window", tabName = "window", icon = icon("window-restore")),
      menuItem("Help", tabName = "help", icon = icon("question-circle"))
    )
  ),
  
  dashboardBody(
    tags$head(
      tags$link(rel = "stylesheet", type = "text/css", href = "https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css"),
      tags$style(HTML(paste0("
        .main-header .logo {
          font-weight: bold;
          font-size: 18px;
          background-color: ", spss_blue, ";
          color: white;
        }
        .content-wrapper, .right-side {
          background-color: white;
        }
        .box {
          box-shadow: 0 1px 3px rgba(0,0,0,0.1);
          border-radius: 0px;
          border: 1px solid #ddd;
        }
        .info-box {
          min-height: 80px;
          border-radius: 0px;
        }
        .bg-blue {
          background-color: ", spss_blue, " !important;
          color: white !important;
        }
        .bg-green {
          background-color: #00a65a !important;
          color: white !important;
        }
        .bg-purple {
          background-color: #605ca8 !important;
          color: white !important;
        }
        .nav-tabs-custom {
          box-shadow: 0 1px 3px rgba(0,0,0,0.1);
          border-radius: 0px;
        }
        .analysis-panel {
          background-color: white;
          padding: 15px;
          border-radius: 0px;
          border: 1px solid #ddd;
          margin-bottom: 15px;
        }
        .data-table {
          font-size: 12px;
        }
        .result-output {
          background-color: white;
          padding: 15px;
          border-radius: 0px;
          border: 1px solid #ddd;
          font-family: 'Courier New', monospace;
          font-size: 13px;
          overflow: auto;
          max-height: 500px;
        }
        .sidebar-menu li.active {
          border-left: 4px solid ", spss_blue, ";
        }
        .skin-blue .main-header .navbar {
          background-color: ", spss_dark_blue, ";
        }
        .skin-blue .main-sidebar {
          background-color: ", spss_light_gray, ";
        }
        .skin-blue .main-sidebar .sidebar .sidebar-menu a {
          color: #444;
        }
        .skin-blue .main-sidebar .sidebar .sidebar-menu a:hover {
          color: ", spss_blue, ";
        }
      ")))
    ),
    
    tabItems(
      # Data Manager Tab
      tabItem(
        tabName = "open",
        h2("Open Data", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Import Data",
            fileInput("file", "Select File",
                      accept = c(".csv", ".xlsx", ".xls", ".sav", ".dta", ".sas7bdat", ".rdata", ".rda"),
                      buttonLabel = "Browse...",
                      placeholder = "No file selected"),
            conditionalPanel(
              condition = "input.file != null",
              selectInput("sep", "Separator (CSV only)", 
                          choices = c(Comma = ",", Semicolon = ";", Tab = "\t", Space = " "), 
                          selected = ","),
              checkboxInput("header", "Header", TRUE),
              actionButton("import_btn", "Open", class = "btn-success", icon = icon("folder-open"))
            ),
            hr(),
            h4("Example Datasets"),
            selectInput("example_data", NULL,
                        choices = c("", "mtcars", "iris", "ToothGrowth", "airquality", "USArrests")),
            actionButton("load_example", "Open Example", class = "btn-info", icon = icon("database"))
          ),
          box(
            width = 8, status = "info", title = "Dataset Information",
            div(class = "analysis-panel",
                verbatimTextOutput("dataset_info"),
                hr(),
                h4("Variable List"),
                withSpinner(DTOutput("var_list"))
            )
          )
        )
      ),
      
      # Data Viewer Tab
      tabItem(
        tabName = "view",
        h2("Data Viewer", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 12, status = "primary", title = "Data Editor",
            div(class = "data-table",
                withSpinner(DTOutput("data_preview"))
            )
          )
        )
      ),
      
      # Compute Variable Tab
      tabItem(
        tabName = "compute",
        h2("Compute Variable", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Compute Variable",
            textInput("target_var", "Target Variable", placeholder = "Enter new variable name"),
            textAreaInput("numeric_expr", "Numeric Expression", 
                          placeholder = "e.g., var1 + var2, or var1 > 5", 
                          rows = 3),
            actionButton("compute_var", "OK", class = "btn-primary", icon = icon("check"))
          ),
          box(
            width = 8, status = "info", title = "Function Reference",
            div(class = "analysis-panel",
                h4("Arithmetic Operators"),
                p("+ Addition, - Subtraction, * Multiplication, / Division, ** Exponentiation"),
                h4("Functions"),
                p("MEAN(), SUM(), SD(), VARIANCE(), MAX(), MIN(), ABS(), SQRT(), LG10(), LN(), EXP(), RND(), TRUNC()")
            )
          )
        )
      ),
      
      # Descriptive Statistics Tab
      tabItem(
        tabName = "descriptives",
        h2("Descriptives", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Variable Selection",
            selectizeInput("desc_vars", "Select Variables", choices = NULL, multiple = TRUE),
            actionButton("run_descriptives", "OK", class = "btn-primary", icon = icon("check")),
            hr(),
            h4("Options"),
            checkboxGroupInput("desc_stats", "Display",
                               choices = c("Mean", "Std. deviation", "Minimum", "Maximum", "Variance", 
                                           "Range", "S.E. mean", "Kurtosis", "Skewness"),
                               selected = c("Mean", "Std. deviation", "Minimum", "Maximum"))
          ),
          box(
            width = 8, status = "info", title = "Descriptive Statistics",
            div(class = "result-output",
                withSpinner(verbatimTextOutput("desc_output"))
            )
          )
        )
      ),
      
      # Frequencies Tab
      tabItem(
        tabName = "frequencies",
        h2("Frequencies", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Variable Selection",
            selectizeInput("freq_vars", "Select Variables", choices = NULL, multiple = TRUE),
            actionButton("run_frequencies", "OK", class = "btn-primary", icon = icon("check")),
            hr(),
            h4("Statistics"),
            checkboxGroupInput("freq_stats", "Display",
                               choices = c("Mean", "Median", "Mode", "Sum", "Std. deviation", 
                                           "Variance", "Range", "Minimum", "Maximum", 
                                           "Skewness", "Kurtosis"),
                               selected = NULL),
            h4("Charts"),
            checkboxGroupInput("freq_charts", "Chart Type",
                               choices = c("Bar charts", "Pie charts", "Histograms"))
          ),
          box(
            width = 8, status = "info", title = "Results",
            tabBox(
              width = 12,
              tabPanel("Statistics", div(class = "result-output", withSpinner(verbatimTextOutput("freq_output")))),
              tabPanel("Charts", withSpinner(plotlyOutput("freq_chart_output", height = "400px")))
            )
          )
        )
      ),
      
      # T-Tests Tab
      tabItem(
        tabName = "independent_ttest",
        h2("Independent Samples T-Test", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Test Variables",
            selectizeInput("ttest_vars", "Test Variable(s)", choices = NULL, multiple = TRUE),
            selectInput("ttest_group_var", "Grouping Variable", choices = NULL),
            textInput("group_def", "Define Groups", placeholder = "e.g., 1,2 or A,B"),
            actionButton("run_ttest", "OK", class = "btn-primary", icon = icon("check")),
            hr(),
            h4("Options"),
            numericInput("ttest_conf", "Confidence Level", value = 95, min = 50, max = 99, step = 1)
          ),
          box(
            width = 8, status = "info", title = "Independent Samples Test",
            div(class = "result-output",
                withSpinner(verbatimTextOutput("ttest_output"))
            ),
            withSpinner(plotlyOutput("ttest_plot", height = "300px"))
          )
        )
      ),
      
      # ANOVA Tab
      tabItem(
        tabName = "anova_menu",
        h2("One-Way ANOVA", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Variables",
            selectInput("anova_dv", "Dependent List", choices = NULL, multiple = TRUE),
            selectInput("anova_factor", "Factor", choices = NULL),
            actionButton("run_anova", "OK", class = "btn-primary", icon = icon("check")),
            hr(),
            h4("Post Hoc"),
            checkboxGroupInput("posthoc_tests", "Equal Variances Assumed",
                               choices = c("LSD", "Bonferroni", "Sidak", "Scheffe", "Tukey", "Duncan")),
            h4("Options"),
            checkboxGroupInput("anova_options", "Statistics",
                               choices = c("Descriptive", "Homogeneity of variance test", "Brown-Forsythe", "Welch"))
          ),
          box(
            width = 8, status = "info", title = "ANOVA",
            tabBox(
              width = 12,
              tabPanel("ANOVA", div(class = "result-output", withSpinner(verbatimTextOutput("anova_output")))),
              tabPanel("Post Hoc", div(class = "result-output", withSpinner(verbatimTextOutput("posthoc_output")))),
              tabPanel("Descriptives", div(class = "result-output", withSpinner(verbatimTextOutput("descriptives_output")))),
              tabPanel("Plots", withSpinner(plotlyOutput("anova_plot", height = "400px")))
            )
          )
        )
      ),
      
      # Correlation Tab
      tabItem(
        tabName = "bivariate",
        h2("Bivariate Correlations", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Variables",
            selectizeInput("cor_vars", "Variables", choices = NULL, multiple = TRUE),
            selectInput("cor_method", "Correlation Coefficients",
                        choices = c("Pearson" = "pearson", "Kendall's tau-b" = "kendall", "Spearman" = "spearman")),
            actionButton("run_correlation", "OK", class = "btn-primary", icon = icon("check")),
            hr(),
            h4("Test of Significance"),
            radioButtons("sig_test", NULL,
                         choices = c("Two-tailed", "One-tailed"), selected = "Two-tailed"),
            checkboxInput("flag_sig", "Flag significant correlations", value = TRUE)
          ),
          box(
            width = 8, status = "info", title = "Correlations",
            div(class = "result-output",
                withSpinner(verbatimTextOutput("cor_output"))
            ),
            withSpinner(plotlyOutput("cor_plot", height = "400px"))
          )
        )
      ),
      
      # Regression Tab
      tabItem(
        tabName = "linear_reg",
        h2("Linear Regression", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Variables",
            selectInput("reg_dep", "Dependent", choices = NULL),
            selectizeInput("reg_indep", "Independent(s)", choices = NULL, multiple = TRUE),
            actionButton("run_regression", "OK", class = "btn-primary", icon = icon("check")),
            hr(),
            h4("Method"),
            selectInput("reg_method", "Method",
                        choices = c("Enter", "Stepwise", "Remove", "Backward", "Forward")),
            h4("Statistics"),
            checkboxGroupInput("reg_stats", "Regression Coefficients",
                               choices = c("Estimates", "Confidence intervals", "Covariance matrix",
                                           "Model fit", "R squared change", "Descriptives",
                                           "Part and partial correlations", "Collinearity diagnostics"))
          ),
          box(
            width = 8, status = "info", title = "Regression",
            tabBox(
              width = 12,
              tabPanel("Model Summary", div(class = "result-output", withSpinner(verbatimTextOutput("reg_summary")))),
              tabPanel("ANOVA", div(class = "result-output", withSpinner(verbatimTextOutput("reg_anova")))),
              tabPanel("Coefficients", div(class = "result-output", withSpinner(verbatimTextOutput("reg_coef")))),
              tabPanel("Residuals", div(class = "result-output", withSpinner(verbatimTextOutput("reg_residuals")))),
              tabPanel("Plots", withSpinner(plotlyOutput("reg_plot", height = "400px")))
            )
          )
        )
      ),
      
      # Factor Analysis Tab
      tabItem(
        tabName = "factor",
        h2("Factor Analysis", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Variables",
            selectizeInput("factor_vars", "Variables", choices = NULL, multiple = TRUE),
            actionButton("run_factor", "OK", class = "btn-primary", icon = icon("check")),
            hr(),
            h4("Desriptives"),
            checkboxGroupInput("factor_descriptives", NULL,
                               choices = c("Initial solution", "KMO and Bartlett's test of sphericity",
                                           "Reproduced", "Anti-image")),
            h4("Extraction"),
            selectInput("factor_method", "Method",
                        choices = c("Principal components", "Unweighted least squares",
                                    "Generalized least squares", "Maximum likelihood",
                                    "Principal axis factoring", "Alpha factoring", "Image factoring")),
            numericInput("factor_n", "Number of factors to extract", value = 1, min = 1, max = 20, step = 1),
            h4("Rotation"),
            selectInput("factor_rotation", "Method",
                        choices = c("None", "Varimax", "Quartimax", "Equamax", "Direct Oblimin", "Promax")),
            h4("Scores"),
            checkboxInput("factor_scores", "Save as variables", value = FALSE),
            checkboxInput("factor_display", "Display factor score coefficient matrix", value = FALSE)
          ),
          box(
            width = 8, status = "info", title = "Factor Analysis",
            tabBox(
              width = 12,
              tabPanel("Communalities", div(class = "result-output", withSpinner(verbatimTextOutput("factor_comm")))),
              tabPanel("Total Variance", div(class = "result-output", withSpinner(verbatimTextOutput("factor_variance")))),
              tabPanel("Component Matrix", div(class = "result-output", withSpinner(verbatimTextOutput("factor_matrix")))),
              tabPanel("Rotated Matrix", div(class = "result-output", withSpinner(verbatimTextOutput("factor_rotated")))),
              tabPanel("Plots", withSpinner(plotlyOutput("factor_plot", height = "400px")))
            )
          )
        )
      ),
      
      # Chart Builder Tab
      tabItem(
        tabName = "chart_builder",
        h2("Chart Builder", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 3, status = "primary", title = "Variables",
            selectizeInput("chart_vars", "Variables", choices = NULL, multiple = TRUE),
            selectInput("chart_group", "Grouping Variable", choices = NULL),
            hr(),
            h4("Chart Type"),
            selectInput("chart_type", "Choose a chart type",
                        choices = c("Bar", "Line", "Area", "Pie", "Scatterplot", "Histogram",
                                    "Boxplot", "Dual Axes", "P-P Plot", "Q-Q Plot")),
            actionButton("generate_chart", "OK", class = "btn-primary", icon = icon("check"))
          ),
          box(
            width = 9, status = "info", title = "Chart Preview",
            withSpinner(plotlyOutput("chart_output", height = "500px"))
          )
        )
      )
    )
  )
)

# Server Logic
server <- function(input, output, session) {
  
  # Reactive data storage
  data <- reactiveVal()
  var_types <- reactiveVal()
  
  # Load example datasets
  observeEvent(input$load_example, {
    req(input$example_data != "")
    
    df <- switch(input$example_data,
                 "mtcars" = mtcars,
                 "iris" = iris,
                 "ToothGrowth" = ToothGrowth,
                 "airquality" = airquality,
                 "USArrests" = USArrests)
    
    data(df)
    
    # Determine variable types
    types <- sapply(df, function(x) {
      if (is.numeric(x)) "Numeric"
      else if (is.factor(x) | is.character(x)) "Categorical"
      else "Other"
    })
    
    var_types(data.frame(
      Variable = names(df),
      Type = types,
      stringsAsFactors = FALSE
    ))
    
    # Update all select inputs
    updateSelectInput(session, "freq_vars", choices = names(df))
    updateSelectInput(session, "desc_vars", choices = names(df))
    updateSelectInput(session, "ttest_vars", choices = names(df))
    updateSelectInput(session, "ttest_group_var", choices = names(df)[types == "Categorical"])
    updateSelectInput(session, "anova_dv", choices = names(df)[types == "Numeric"])
    updateSelectInput(session, "anova_factor", choices = names(df)[types == "Categorical"])
    updateSelectInput(session, "cor_vars", choices = names(df)[types == "Numeric"])
    updateSelectInput(session, "reg_dep", choices = names(df)[types == "Numeric"])
    updateSelectInput(session, "reg_indep", choices = names(df))
    updateSelectInput(session, "factor_vars", choices = names(df)[types == "Numeric"])
    updateSelectInput(session, "chart_vars", choices = names(df))
    updateSelectInput(session, "chart_group", choices = c("None", names(df)[types == "Categorical"]))
  })
  
  # Data import and processing
  observeEvent(input$import_btn, {
    req(input$file)
    
    ext <- tools::file_ext(input$file$name)
    
    tryCatch({
      df <- switch(ext,
                   csv = read.csv(input$file$datapath, 
                                  header = input$header, 
                                  sep = input$sep),
                   xlsx = read_excel(input$file$datapath),
                   xls = read_excel(input$file$datapath),
                   sav = read_sav(input$file$datapath),
                   dta = read_dta(input$file$datapath),
                   sas7bdat = read_sas(input$file$datapath),
                   rdata = load(input$file$datapath),
                   rda = load(input$file$datapath),
                   stop("Invalid file type")
      )
      
      # Convert to data frame if necessary
      if (!is.data.frame(df)) {
        df <- as.data.frame(df)
      }
      
      data(df)
      
      # Determine variable types
      types <- sapply(df, function(x) {
        if (is.numeric(x)) "Numeric"
        else if (is.factor(x) | is.character(x)) "Categorical"
        else "Other"
      })
      
      var_types(data.frame(
        Variable = names(df),
        Type = types,
        stringsAsFactors = FALSE
      ))
      
      # Update all select inputs
      updateSelectInput(session, "freq_vars", choices = names(df))
      updateSelectInput(session, "desc_vars", choices = names(df))
      updateSelectInput(session, "ttest_vars", choices = names(df))
      updateSelectInput(session, "ttest_group_var", choices = names(df)[types == "Categorical"])
      updateSelectInput(session, "anova_dv", choices = names(df)[types == "Numeric"])
      updateSelectInput(session, "anova_factor", choices = names(df)[types == "Categorical"])
      updateSelectInput(session, "cor_vars", choices = names(df)[types == "Numeric"])
      updateSelectInput(session, "reg_dep", choices = names(df)[types == "Numeric"])
      updateSelectInput(session, "reg_indep", choices = names(df))
      updateSelectInput(session, "factor_vars", choices = names(df)[types == "Numeric"])
      updateSelectInput(session, "chart_vars", choices = names(df))
      updateSelectInput(session, "chart_group", choices = c("None", names(df)[types == "Categorical"]))
      
    }, error = function(e) {
      showNotification(paste("Error:", e$message), type = "error")
    })
  })
  
  # Dataset information
  output$dataset_info <- renderPrint({
    req(data())
    df <- data()
    cat("Dataset: ", ifelse(is.null(input$file), input$example_data, input$file$name), "\n")
    cat("Number of cases: ", nrow(df), "\n")
    cat("Number of variables: ", ncol(df), "\n")
    cat("\nVariable types:\n")
    print(table(sapply(df, class)))
  })
  
  # Variable list
  output$var_list <- renderDT({
    req(var_types())
    datatable(var_types(), options = list(dom = 't', pageLength = 10), rownames = FALSE)
  })
  
  # Data preview
  output$data_preview <- renderDT({
    req(data())
    datatable(data(), options = list(scrollX = TRUE, pageLength = 10))
  })
  
  # Descriptive statistics
  output$desc_output <- renderPrint({
    req(data(), input$desc_vars)
    
    df <- data()
    vars <- df[, input$desc_vars, drop = FALSE]
    
    if (ncol(vars) == 0) return("Please select at least one variable")
    
    # Calculate requested statistics
    results <- list()
    
    if ("Mean" %in% input$desc_stats) {
      means <- sapply(vars, mean, na.rm = TRUE)
      results$Mean <- means
    }
    
    if ("Std. deviation" %in% input$desc_stats) {
      sds <- sapply(vars, sd, na.rm = TRUE)
      results$`Std. Deviation` <- sds
    }
    
    if ("Minimum" %in% input$desc_stats) {
      mins <- sapply(vars, min, na.rm = TRUE)
      results$Minimum <- mins
    }
    
    if ("Maximum" %in% input$desc_stats) {
      maxs <- sapply(vars, max, na.rm = TRUE)
      results$Maximum <- maxs
    }
    
    if ("Variance" %in% input$desc_stats) {
      variances <- sapply(vars, var, na.rm = TRUE)
      results$Variance <- variances
    }
    
    if ("Range" %in% input$desc_stats) {
      ranges <- sapply(vars, function(x) diff(range(x, na.rm = TRUE)))
      results$Range <- ranges
    }
    
    if ("S.E. mean" %in% input$desc_stats) {
      se_mean <- sapply(vars, function(x) sd(x, na.rm = TRUE)/sqrt(length(na.omit(x))))
      results$`Std. Error Mean` <- se_mean
    }
    
    if ("Kurtosis" %in% input$desc_stats) {
      kurts <- sapply(vars, psych::kurtosi, na.rm = TRUE)
      results$Kurtosis <- kurts
    }
    
    if ("Skewness" %in% input$desc_stats) {
      skews <- sapply(vars, psych::skew, na.rm = TRUE)
      results$Skewness <- skews
    }
    
    # Format output in SPSS style
    cat("Descriptive Statistics\n")
    cat("======================\n\n")
    
    # Create a table similar to SPSS output
    result_matrix <- matrix(nrow = length(input$desc_vars), ncol = length(names(results)))
    colnames(result_matrix) <- names(results)
    rownames(result_matrix) <- input$desc_vars
    
    for (i in 1:length(names(results))) {
      result_matrix[, i] <- results[[i]]
    }
    
    print(round(result_matrix, 3))
    
    cat("\nN =", nrow(df), "\n")
  })
  
  # Frequency tables
  output$freq_output <- renderPrint({
    req(data(), input$freq_vars)
    
    if (length(input$freq_vars) == 0) return("Please select at least one variable")
    
    cat("Frequencies\n")
    cat("===========\n\n")
    
    for (var_name in input$freq_vars) {
      var <- data()[[var_name]]
      
      freq_table <- table(var)
      prop_table <- prop.table(freq_table) * 100
      cum_table <- cumsum(prop_table)
      
      cat("Variable: ", var_name, "\n\n")
      
      result_df <- data.frame(
        Frequency = as.numeric(freq_table),
        Percent = round(as.numeric(prop_table), 1),
        ValidPercent = round(as.numeric(prop_table), 1),
        CumulativePercent = round(as.numeric(cum_table), 1)
      )
      
      # Add value labels if available
      if (is.factor(var)) {
        result_df <- cbind(Value = levels(var), result_df)
      } else {
        result_df <- cbind(Value = names(freq_table), result_df)
      }
      
      print(result_df, row.names = FALSE)
      cat("\n")
      
      # Statistics if requested
      if (length(input$freq_stats) > 0 && is.numeric(var)) {
        cat("Statistics for ", var_name, "\n")
        cat("----------------------------\n")
        
        stats_list <- list()
        if ("Mean" %in% input$freq_stats) stats_list$Mean <- mean(var, na.rm = TRUE)
        if ("Median" %in% input$freq_stats) stats_list$Median <- median(var, na.rm = TRUE)
        if ("Std. deviation" %in% input$freq_stats) stats_list$`Std. Deviation` <- sd(var, na.rm = TRUE)
        if ("Variance" %in% input$freq_stats) stats_list$Variance <- var(var, na.rm = TRUE)
        if ("Range" %in% input$freq_stats) stats_list$Range <- diff(range(var, na.rm = TRUE))
        if ("Minimum" %in% input$freq_stats) stats_list$Minimum <- min(var, na.rm = TRUE)
        if ("Maximum" %in% input$freq_stats) stats_list$Maximum <- max(var, na.rm = TRUE)
        if ("Sum" %in% input$freq_stats) stats_list$Sum <- sum(var, na.rm = TRUE)
        if ("Skewness" %in% input$freq_stats) stats_list$Skewness <- psych::skew(var, na.rm = TRUE)
        if ("Kurtosis" %in% input$freq_stats) stats_list$Kurtosis <- psych::kurtosi(var, na.rm = TRUE)
        
        for (stat in names(stats_list)) {
          cat(stat, ": ", round(stats_list[[stat]], 3), "\n")
        }
        cat("\n")
      }
    }
  })
  
  # Frequency charts
  output$freq_chart_output <- renderPlotly({
    req(data(), input$freq_vars)
    
    if (length(input$freq_vars) == 0) return()
    
    var_name <- input$freq_vars[1]
    var <- data()[[var_name]]
    
    if ("Bar charts" %in% input$freq_charts) {
      p <- ggplot(data(), aes_string(x = var_name)) +
        geom_bar(fill = spss_blue, alpha = 0.7) +
        theme_minimal() +
        labs(title = paste("Bar Chart of", var_name))
    } else if ("Pie charts" %in% input$freq_charts) {
      freq <- table(var)
      p <- plot_ly(labels = names(freq), values = as.numeric(freq), type = "pie")
    } else if ("Histograms" %in% input$freq_charts && is.numeric(var)) {
      p <- ggplot(data(), aes_string(x = var_name)) +
        geom_histogram(fill = spss_blue, alpha = 0.7, bins = 30) +
        theme_minimal() +
        labs(title = paste("Histogram of", var_name))
    } else {
      return()
    }
    
    if ("Bar charts" %in% input$freq_charts || "Histograms" %in% input$freq_charts) {
      ggplotly(p)
    } else {
      p
    }
  })
  
  # T-Test output
  output$ttest_output <- renderPrint({
    req(data(), input$ttest_vars, input$ttest_group_var, input$group_def)
    
    # Parse group definition
    groups <- trimws(unlist(strsplit(input$group_def, ",")))
    if (length(groups) != 2) {
      return("Please define exactly two groups separated by a comma (e.g., '1,2' or 'A,B')")
    }
    
    # Filter data based on groups
    group_var <- data()[[input$ttest_group_var]]
    if (!all(groups %in% unique(group_var))) {
      return("One or both group values not found in the grouping variable")
    }
    
    # Create subset for each group
    group1_data <- data()[group_var == groups[1], input$ttest_vars, drop = FALSE]
    group2_data <- data()[group_var == groups[2], input$ttest_vars, drop = FALSE]
    
    # Perform t-test for each variable
    conf_level <- input$ttest_conf / 100
    
    cat("Independent Samples Test\n")
    cat("========================\n\n")
    
    for (var in input$ttest_vars) {
      cat("Variable: ", var, "\n")
      cat("------------------------\n")
      
      # Check if variable is numeric
      if (!is.numeric(data()[[var]])) {
        cat("Skipping non-numeric variable:", var, "\n\n")
        next
      }
      
      # Perform t-test
      test_result <- t.test(
        group1_data[[var]], 
        group2_data[[var]], 
        conf.level = conf_level
      )
      
      # Format output similar to SPSS
      cat("t-value:", round(test_result$statistic, 3), "\n")
      cat("df:", round(test_result$parameter, 1), "\n")
      cat("p-value:", format.pval(test_result$p.value, digits = 3), "\n")
      cat("Mean Difference:", round(diff(test_result$estimate), 3), "\n")
      cat("95% CI for Difference: [", 
          round(test_result$conf.int[1], 3), ", ", 
          round(test_result$conf.int[2], 3), "]\n\n")
    }
  })
  
  # T-Test plot
  output$ttest_plot <- renderPlotly({
    req(data(), input$ttest_vars, input$ttest_group_var, input$group_def)
    
    # Parse group definition
    groups <- trimws(unlist(strsplit(input$group_def, ",")))
    if (length(groups) != 2) return()
    
    # Use the first variable for plotting
    var <- input$ttest_vars[1]
    
    # Filter data for the two groups
    filtered_data <- data()[data()[[input$ttest_group_var]] %in% groups, ]
    
    p <- ggplot(filtered_data, aes_string(x = input$ttest_group_var, y = var, fill = input$ttest_group_var)) +
      geom_boxplot(alpha = 0.7) +
      scale_fill_manual(values = c(spss_blue, "#E27D72")) +
      theme_minimal() +
      labs(title = paste("Boxplot of", var, "by", input$ttest_group_var))
    
    ggplotly(p)
  })
  
  # ANOVA output
  output$anova_output <- renderPrint({
    req(data(), input$anova_dv, input$anova_factor)
    
    formula <- as.formula(paste(input$anova_dv, "~", input$anova_factor))
    anova_result <- aov(formula, data = data())
    
    cat("ANOVA\n")
    cat("=====\n\n")
    
    summary(anova_result)
  })
  
  # Correlation output
  output$cor_output <- renderPrint({
    req(data(), input$cor_vars)
    
    if (length(input$cor_vars) < 2) return("Please select at least two variables")
    
    cor_data <- data()[, input$cor_vars, drop = FALSE]
    cor_matrix <- cor(cor_data, use = "complete.obs", method = input$cor_method)
    
    cat("Correlations\n")
    cat("============\n\n")
    
    # Print correlation matrix
    print(round(cor_matrix, 3))
    
    # Add significance stars if requested
    if (input$flag_sig) {
      cat("\n* Correlation is significant at the 0.05 level (2-tailed).\n")
      cat("** Correlation is significant at the 0.01 level (2-tailed).\n")
    }
  })
  
  # Correlation plot
  output$cor_plot <- renderPlotly({
    req(data(), input$cor_vars)
    
    if (length(input$cor_vars) < 2) return()
    
    cor_data <- data()[, input$cor_vars, drop = FALSE]
    cor_matrix <- cor(cor_data, use = "complete.obs", method = input$cor_method)
    
    # Create a heatmap
    plot_ly(
      x = colnames(cor_matrix), 
      y = rownames(cor_matrix),
      z = cor_matrix,
      type = "heatmap",
      colors = colorRamp(c(spss_blue, "white", "#E27D72")),
      zmin = -1, zmax = 1
    ) %>%
      layout(
        title = "Correlation Matrix",
        xaxis = list(title = ""),
        yaxis = list(title = "")
      )
  })
  
  # Regression output
  output$reg_summary <- renderPrint({
    req(data(), input$reg_dep, input$reg_indep)
    
    if (length(input$reg_indep) == 0) return("Please select at least one independent variable")
    
    formula <- as.formula(paste(input$reg_dep, "~", paste(input$reg_indep, collapse = "+")))
    model <- lm(formula, data = data())
    
    cat("Regression\n")
    cat("==========\n\n")
    print(summary(model))
  })
  
  # Regression ANOVA
  output$reg_anova <- renderPrint({
    req(data(), input$reg_dep, input$reg_indep)
    
    if (length(input$reg_indep) == 0) return()
    
    formula <- as.formula(paste(input$reg_dep, "~", paste(input$reg_indep, collapse = "+")))
    model <- lm(formula, data = data())
    
    cat("ANOVA\n")
    cat("=====\n\n")
    anova(model)
  })
  
  # Regression coefficients
  output$reg_coef <- renderPrint({
    req(data(), input$reg_dep, input$reg_indep)
    
    if (length(input$reg_indep) == 0) return()
    
    formula <- as.formula(paste(input$reg_dep, "~", paste(input$reg_indep, collapse = "+")))
    model <- lm(formula, data = data())
    
    cat("Coefficients\n")
    cat("============\n\n")
    
    coef_df <- as.data.frame(summary(model)$coefficients)
    colnames(coef_df) <- c("B", "Std. Error", "t", "Sig.")
    print(round(coef_df, 4))
  })
  
  # Chart output
  output$chart_output <- renderPlotly({
    req(data(), input$chart_vars)
    
    if (input$chart_type == "Bar") {
      req(length(input$chart_vars) >= 1)
      var <- input$chart_vars[1]
      
      if (input$chart_group != "None") {
        p <- ggplot(data(), aes_string(x = input$chart_group, fill = input$chart_group)) +
          geom_bar(alpha = 0.7) +
          scale_fill_manual(values = c(spss_blue, "#E27D72", "#68C3A3", "#F7C46C")) +
          theme_minimal() +
          labs(title = paste("Bar Chart of", var))
      } else {
        p <- ggplot(data(), aes_string(x = var)) +
          geom_bar(fill = spss_blue, alpha = 0.7) +
          theme_minimal() +
          labs(title = paste("Bar Chart of", var))
      }
    } else if (input$chart_type == "Histogram") {
      req(length(input$chart_vars) >= 1)
      var <- input$chart_vars[1]
      
      p <- ggplot(data(), aes_string(x = var)) +
        geom_histogram(fill = spss_blue, alpha = 0.7, bins = 30) +
        theme_minimal() +
        labs(title = paste("Histogram of", var))
    } else if (input$chart_type == "Boxplot") {
      req(length(input$chart_vars) >= 1)
      var <- input$chart_vars[1]
      
      if (input$chart_group != "None") {
        p <- ggplot(data(), aes_string(x = input$chart_group, y = var, fill = input$chart_group)) +
          geom_boxplot(alpha = 0.7) +
          scale_fill_manual(values = c(spss_blue, "#E27D72", "#68C3A3", "#F7C46C")) +
          theme_minimal() +
          labs(title = paste("Boxplot of", var, "by", input$chart_group))
      } else {
        p <- ggplot(data(), aes_string(y = var)) +
          geom_boxplot(fill = spss_blue, alpha = 0.7) +
          theme_minimal() +
          labs(title = paste("Boxplot of", var))
      }
    } else if (input$chart_type == "Scatterplot") {
      req(length(input$chart_vars) >= 2)
      x_var <- input$chart_vars[1]
      y_var <- input$chart_vars[2]
      
      if (input$chart_group != "None") {
        p <- ggplot(data(), aes_string(x = x_var, y = y_var, color = input$chart_group)) +
          geom_point(alpha = 0.7) +
          scale_color_manual(values = c(spss_blue, "#E27D72", "#68C3A3", "#F7C46C")) +
          theme_minimal() +
          labs(title = paste("Scatterplot of", y_var, "vs", x_var))
      } else {
        p <- ggplot(data(), aes_string(x = x_var, y = y_var)) +
          geom_point(color = spss_blue, alpha = 0.7) +
          theme_minimal() +
          labs(title = paste("Scatterplot of", y_var, "vs", x_var))
      }
    } else {
      return()
    }
    
    ggplotly(p)
  })
  
  # Compute variable functionality
  observeEvent(input$compute_var, {
    req(data(), input$target_var, input$numeric_expr)
    
    df <- data()
    tryCatch({
      # Parse and evaluate the expression
      expr <- parse(text = input$numeric_expr)
      df[[input$target_var]] <- eval(expr, envir = df)
      data(df)
      
      # Update variable types
      types <- sapply(df, function(x) {
        if (is.numeric(x)) "Numeric"
        else if (is.factor(x) | is.character(x)) "Categorical"
        else "Other"
      })
      
      var_types(data.frame(
        Variable = names(df),
        Type = types,
        stringsAsFactors = FALSE
      ))
      
      # Update all select inputs
      updateSelectInput(session, "freq_vars", choices = names(df))
      updateSelectInput(session, "desc_vars", choices = names(df))
      updateSelectInput(session, "ttest_vars", choices = names(df))
      updateSelectInput(session, "ttest_group_var", choices = names(df)[types == "Categorical"])
      updateSelectInput(session, "anova_dv", choices = names(df)[types == "Numeric"])
      updateSelectInput(session, "anova_factor", choices = names(df)[types == "Categorical"])
      updateSelectInput(session, "cor_vars", choices = names(df)[types == "Numeric"])
      updateSelectInput(session, "reg_dep", choices = names(df)[types == "Numeric"])
      updateSelectInput(session, "reg_indep", choices = names(df))
      updateSelectInput(session, "factor_vars", choices = names(df)[types == "Numeric"])
      updateSelectInput(session, "chart_vars", choices = names(df))
      updateSelectInput(session, "chart_group", choices = c("None", names(df)[types == "Categorical"]))
      
      showNotification(paste("Variable", input$target_var, "created successfully"), type = "message")
    }, error = function(e) {
      showNotification(paste("Error:", e$message), type = "error")
    })
  })
}

# Run the application
shinyApp(ui = ui, server = server)
