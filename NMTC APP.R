# app.R
# SPSS-Style Statistical Analysis Software - Enhanced Version
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
library(MASS)
library(nnet)
library(foreign)
library(rstatix)
#library(flextable)
library(officer)
library(openxlsx)
library(ggpubr)
library(RColorBrewer)
library(lavaan)  # For EFA/CFA
library(semTools) # For McDonald's omega

# SPSS Color Scheme
spss_blue <- "#326EA6"
spss_dark_blue <- "#1A3A5F"
spss_light_gray <- "#F2F2F2"
spss_dark_gray <- "#7F7F7F"

# UI Definition - SPSS Style Interface
ui <- dashboardPage(
  skin = "blue",
  dashboardHeader(
    title = span(icon("chart-bar"), "TNMTC DataLab", style = paste0("color: white; font-weight: bold; font-size: 18px; background-color:", spss_blue)),
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
                        menuSubItem("One-Sample T Test", tabName = "one_sample_ttest"),
                        menuSubItem("Independent Samples T Test", tabName = "independent_ttest"),
                        menuSubItem("Paired Samples T Test", tabName = "paired_ttest"),
                        menuSubItem("ANOVA", tabName = "anova_menu")
               ),
               menuItem("Correlation", tabName = "correlation_menu", icon = icon("random"),
                        menuSubItem("Bivariate", tabName = "bivariate")
               ),
               menuItem("Regression", tabName = "regression_menu", icon = icon("chart-line"),
                        menuSubItem("Linear", tabName = "linear_reg"),
                        menuSubItem("Binary Logistic", tabName = "logistic_reg"),
                        menuSubItem("Ordinal", tabName = "ordinal_reg"),
                        menuSubItem("Multinomial", tabName = "multinomial_reg"),
                        menuSubItem("Poisson", tabName = "poisson_reg")
               ),
               menuItem("Dimension Reduction", tabName = "dim_reduction", icon = icon("compress"),
                        menuSubItem("Factor Analysis", tabName = "factor_analysis"),
                        menuSubItem("Confirmatory FA", tabName = "cfa")
               ),
               menuItem("Scale", tabName = "scale_menu", icon = icon("sliders-h"),
                        menuSubItem("Reliability Analysis", tabName = "reliability")
               ),
               menuItem("Nonparametric Tests", tabName = "nonparametric", icon = icon("check-square"),
                        menuSubItem("Chi-Square", tabName = "chi_square"),
                        menuSubItem("Mann-Whitney U", tabName = "mann_whitney"),
                        menuSubItem("Kruskal-Wallis", tabName = "kruskal_wallis"),
                        menuSubItem("Wilcoxon Signed-Rank", tabName = "wilcoxon")
               )
      ),
      menuItem("Graphs", tabName = "graphs", icon = icon("chart-area"),
               menuSubItem("Chart Builder", tabName = "chart_builder")),
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
        /* FIX FOR SUB-MENU ITEMS - Make background clear like main tabs */
        .skin-blue .main-sidebar .sidebar .sidebar-menu .treeview-menu {
          background-color: #F8F9FA !important;
        }
        .skin-blue .main-sidebar .sidebar .sidebar-menu .treeview-menu li a {
          color: #444 !important;
          background-color: #F8F9FA !important;
        }
        .skin-blue .main-sidebar .sidebar .sidebar-menu .treeview-menu li.active a,
        .skin-blue .main-sidebar .sidebar .sidebar-menu .treeview-menu li a:hover {
          color: ", spss_blue, " !important;
          background-color: #E9ECEF !important;
        }
        .export-btn {
          margin-right: 5px;
          margin-bottom: 5px;
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
              actionButton("import_btn", "Import", class = "btn-success", icon = icon("folder-open"))
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
      
      # Save Data Tab
      tabItem(
        tabName = "save",
        h2("Save Data", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Save Options",
            selectInput("save_format", "Format", 
                        choices = c("CSV", "Excel", "SPSS", "Stata", "RData")),
            textInput("save_filename", "File Name", value = "data"),
            actionButton("save_btn", "Save", class = "btn-primary", icon = icon("save"))
          ),
          box(
            width = 8, status = "info", title = "Data Preview",
            div(class = "data-table",
                withSpinner(DTOutput("save_preview"))
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
      
      # Define Variable Properties Tab
      tabItem(
        tabName = "define_var",
        h2("Define Variable Properties", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Variable Selection",
            selectInput("define_var", "Select Variable", choices = NULL),
            textInput("var_label", "Variable Label", value = ""),
            selectInput("var_type", "Variable Type", 
                        choices = c("Numeric", "Categorical", "Date", "String")),
            conditionalPanel(
              condition = "input.var_type == 'Categorical'",
              textAreaInput("value_labels", "Value Labels", 
                            placeholder = "Format: 1=Male, 2=Female")
            ),
            actionButton("apply_var_props", "Apply", class = "btn-primary", icon = icon("check"))
          ),
          box(
            width = 8, status = "info", title = "Variable Properties",
            div(class = "result-output",
                withSpinner(verbatimTextOutput("var_props_output"))
            )
          )
        )
      ),
      
      # Sort Cases Tab
      tabItem(
        tabName = "sort_cases",
        h2("Sort Cases", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Sort Options",
            selectizeInput("sort_vars", "Sort by", choices = NULL, multiple = TRUE),
            radioButtons("sort_order", "Sort Order", 
                         choices = c("Ascending", "Descending"), selected = "Ascending"),
            actionButton("sort_btn", "Sort", class = "btn-primary", icon = icon("sort"))
          ),
          box(
            width = 8, status = "info", title = "Sorted Data Preview",
            div(class = "data-table",
                withSpinner(DTOutput("sorted_preview"))
            )
          )
        )
      ),
      
      # Select Cases Tab
      tabItem(
        tabName = "select_cases",
        h2("Select Cases", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Selection Criteria",
            radioButtons("select_method", "Method", 
                         choices = c("All cases", "If condition is satisfied", "Random sample")),
            conditionalPanel(
              condition = "input.select_method == 'If condition is satisfied'",
              textAreaInput("select_condition", "Condition", 
                            placeholder = "e.g., age > 30 & gender == 'Male'")
            ),
            conditionalPanel(
              condition = "input.select_method == 'Random sample'",
              numericInput("sample_percent", "Sample Percentage", value = 10, min = 1, max = 100),
              actionButton("sample_btn", "Take Sample", class = "btn-info", icon = icon("random"))
            ),
            actionButton("select_btn", "Apply Selection", class = "btn-primary", icon = icon("filter"))
          ),
          box(
            width = 8, status = "info", title = "Selected Data Preview",
            div(class = "data-table",
                withSpinner(DTOutput("selected_preview"))
            )
          )
        )
      ),
      
      # Weight Cases Tab
      tabItem(
        tabName = "weight_cases",
        h2("Weight Cases", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Weight Options",
            radioButtons("weight_method", "Weight Cases by", 
                         choices = c("Do not weight cases", "Weight cases by")),
            conditionalPanel(
              condition = "input.weight_method == 'Weight cases by'",
              selectInput("weight_var", "Frequency Variable", choices = NULL)
            ),
            actionButton("weight_btn", "Apply Weighting", class = "btn-primary", icon = icon("balance-scale"))
          ),
          box(
            width = 8, status = "info", title = "Weighted Data Information",
            div(class = "result-output",
                withSpinner(verbatimTextOutput("weight_output"))
            )
          )
        )
      ),
      
      # Transform Tab
      tabItem(
        tabName = "transform",
        h2("Data Transformation", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Transformation Options",
            selectInput("transform_type", "Transformation Type",
                        choices = c("Recode", "Compute", "Binning", "Log Transform", "Standardize")),
            conditionalPanel(
              condition = "input.transform_type == 'Recode'",
              selectInput("recode_var", "Variable to Recode", choices = NULL),
              radioButtons("recode_into", "Recode Into",
                           choices = c("Same variable", "Different variable")),
              conditionalPanel(
                condition = "input.recode_into == 'Different variable'",
                textInput("new_recode_var", "New Variable Name", value = "")
              ),
              textAreaInput("recode_rules", "Recode Rules",
                            placeholder = "Format: 1=5, 2=4, 3=3, 4=2, 5=1")
            ),
            conditionalPanel(
              condition = "input.transform_type == 'Compute'",
              textInput("compute_var", "Target Variable", placeholder = "Enter new variable name"),
              textAreaInput("compute_expr", "Expression", 
                            placeholder = "e.g., var1 + var2, or var1 > 5", 
                            rows = 3)
            ),
            conditionalPanel(
              condition = "input.transform_type == 'Binning'",
              selectInput("bin_var", "Variable to Bin", choices = NULL),
              textInput("binned_var", "New Variable Name", value = ""),
              numericInput("num_bins", "Number of Bins", value = 4, min = 2, max = 20)
            ),
            conditionalPanel(
              condition = "input.transform_type == 'Log Transform'",
              selectInput("log_var", "Variable to Transform", choices = NULL),
              radioButtons("log_type", "Log Type",
                           choices = c("Natural Log", "Log10")),
              textInput("log_var_name", "New Variable Name", value = "")
            ),
            conditionalPanel(
              condition = "input.transform_type == 'Standardize'",
              selectInput("standardize_var", "Variable to Standardize", choices = NULL),
              textInput("z_var_name", "New Variable Name", value = "")
            ),
            actionButton("apply_transform", "Apply Transformation", class = "btn-primary", icon = icon("check"))
          ),
          box(
            width = 8, status = "info", title = "Transformed Data Preview",
            div(class = "data-table",
                withSpinner(DTOutput("transform_preview"))
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
                                           "Range", "S.E. mean", "Kurtosis", "Skewness", "Valid N", "Missing N"),
                               selected = c("Mean", "Std. deviation", "Minimum", "Maximum")),
            hr(),
            h4("Export Options"),
            downloadButton("export_descriptives_word", "Word", class = "export-btn"),
            downloadButton("export_descriptives_excel", "Excel", class = "export-btn")
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
                               choices = c("Bar charts", "Pie charts", "Histograms")),
            hr(),
            h4("Export Options"),
            downloadButton("export_frequencies_word", "Word", class = "export-btn"),
            downloadButton("export_frequencies_excel", "Excel", class = "export-btn")
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
      
      # Explore Tab
      tabItem(
        tabName = "explore",
        h2("Explore", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Variable Selection",
            selectizeInput("explore_vars", "Dependent Variables", choices = NULL, multiple = TRUE),
            selectInput("explore_factor", "Factor Variable", choices = NULL),
            actionButton("run_explore", "OK", class = "btn-primary", icon = icon("check")),
            hr(),
            h4("Statistics"),
            checkboxGroupInput("explore_stats", "Display",
                               choices = c("Descriptives", "M-estimators", "Outliers", "Normality tests")),
            h4("Plots"),
            checkboxGroupInput("explore_plots", "Plot Type",
                               choices = c("Boxplots", "Histograms", "Normality plots")),
            hr(),
            h4("Export Options"),
            downloadButton("export_explore_word", "Word", class = "export-btn"),
            downloadButton("export_explore_excel", "Excel", class = "export-btn")
          ),
          box(
            width = 8, status = "info", title = "Explore Results",
            tabBox(
              width = 12,
              tabPanel("Statistics", div(class = "result-output", withSpinner(verbatimTextOutput("explore_output")))),
              tabPanel("Plots", withSpinner(plotlyOutput("explore_plot_output", height = "500px")))
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
            numericInput("ttest_conf", "Confidence Level", value = 95, min = 50, max = 99, step = 1),
            radioButtons("ttest_tails", "Test Type",
                         choices = c("Two-tailed", "One-tailed"), selected = "Two-tailed"),
            hr(),
            h4("Export Options"),
            downloadButton("export_ttest_word", "Word", class = "export-btn"),
            downloadButton("export_ttest_excel", "Excel", class = "export-btn")
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
      
      # One Sample T-Test Tab
      tabItem(
        tabName = "one_sample_ttest",
        h2("One Sample T-Test", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Test Variables",
            selectizeInput("onesample_vars", "Test Variable(s)", choices = NULL, multiple = TRUE),
            numericInput("test_value", "Test Value", value = 0),
            actionButton("run_onesample", "OK", class = "btn-primary", icon = icon("check")),
            hr(),
            h4("Options"),
            numericInput("onesample_conf", "Confidence Level", value = 95, min = 50, max = 99, step = 1),
            radioButtons("onesample_tails", "Test Type",
                         choices = c("Two-tailed", "One-tailed"), selected = "Two-tailed"),
            hr(),
            h4("Export Options"),
            downloadButton("export_onesample_word", "Word", class = "export-btn"),
            downloadButton("export_onesample_excel", "Excel", class = "export-btn")
          ),
          box(
            width = 8, status = "info", title = "One Sample T-Test",
            div(class = "result-output",
                withSpinner(verbatimTextOutput("onesample_output"))
            )
          )
        )
      ),
      
      # Paired T-Test Tab
      tabItem(
        tabName = "paired_ttest",
        h2("Paired Samples T-Test", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Test Variables",
            selectizeInput("paired_vars", "Paired Variables", choices = NULL, multiple = TRUE),
            actionButton("run_paired", "OK", class = "btn-primary", icon = icon("check")),
            hr(),
            h4("Options"),
            numericInput("paired_conf", "Confidence Level", value = 95, min = 50, max = 99, step = 1),
            radioButtons("paired_tails", "Test Type",
                         choices = c("Two-tailed", "One-tailed"), selected = "Two-tailed"),
            hr(),
            h4("Export Options"),
            downloadButton("export_paired_word", "Word", class = "export-btn"),
            downloadButton("export_paired_excel", "Excel", class = "export-btn")
          ),
          box(
            width = 8, status = "info", title = "Paired Samples Test",
            div(class = "result-output",
                withSpinner(verbatimTextOutput("paired_output"))
            ),
            withSpinner(plotlyOutput("paired_plot", height = "300px"))
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
                               choices = c("Descriptive", "Homogeneity of variance test", "Brown-Forsythe", "Welch")),
            hr(),
            h4("Export Options"),
            downloadButton("export_anova_word", "Word", class = "export-btn"),
            downloadButton("export_anova_excel", "Excel", class = "export-btn")
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
      
      # Two-Way ANOVA Tab
      tabItem(
        tabName = "twoway_anova",
        h2("Two-Way ANOVA", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Variables",
            selectInput("twoway_dv", "Dependent Variable", choices = NULL),
            selectInput("twoway_factor1", "Factor 1", choices = NULL),
            selectInput("twoway_factor2", "Factor 2", choices = NULL),
            actionButton("run_twoway", "OK", class = "btn-primary", icon = icon("check")),
            hr(),
            h4("Options"),
            checkboxGroupInput("twoway_options", "Display",
                               choices = c("Descriptive statistics", "Homogeneity tests", "Effect size", "Observed power")),
            h4("Plots"),
            checkboxGroupInput("twoway_plots", "Plot Type",
                               choices = c("Interaction plot", "Profile plot")),
            hr(),
            h4("Export Options"),
            downloadButton("export_twoway_word", "Word", class = "export-btn"),
            downloadButton("export_twoway_excel", "Excel", class = "export-btn")
          ),
          box(
            width = 8, status = "info", title = "Two-Way ANOVA",
            tabBox(
              width = 12,
              tabPanel("ANOVA", div(class = "result-output", withSpinner(verbatimTextOutput("twoway_output")))),
              tabPanel("Descriptives", div(class = "result-output", withSpinner(verbatimTextOutput("twoway_descriptives")))),
              tabPanel("Plots", withSpinner(plotlyOutput("twoway_plot", height = "400px")))
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
            checkboxInput("flag_sig", "Flag significant correlations", value = TRUE),
            hr(),
            h4("Export Options"),
            downloadButton("export_correlation_word", "Word", class = "export-btn"),
            downloadButton("export_correlation_excel", "Excel", class = "export-btn")
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
      
      # Linear Regression Tab
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
                                           "Part and partial correlations", "Collinearity diagnostics")),
            h4("Plots"),
            checkboxGroupInput("reg_plots", "Plot Type",
                               choices = c("Residuals vs Fitted", "Normal Q-Q", "Scale-Location", "Residuals vs Leverage")),
            hr(),
            h4("Export Options"),
            downloadButton("export_regression_word", "Word", class = "export-btn"),
            downloadButton("export_regression_excel", "Excel", class = "export-btn")
          ),
          box(
            width = 8, status = "info", title = "Regression",
            tabBox(
              width = 12,
              tabPanel("Model Summary", div(class = "result-output", withSpinner(verbatimTextOutput("reg_summary")))),
              tabPanel("ANOVA", div(class = "result-output", withSpinner(verbatimTextOutput("reg_anova")))),
              tabPanel("Coefficients", div(class = "result-output", withSpinner(verbatimTextOutput("reg_coef")))),
              tabPanel("Residuals", div(class = "result-output", withSpinner(verbatimTextOutput("reg_residuals")))),
              tabPanel("Plots", withSpinner(plotOutput("reg_plot", height = "500px")))
            )
          )
        )
      ),
      
      # Logistic Regression Tab
      tabItem(
        tabName = "logistic_reg",
        h2("Binary Logistic Regression", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Variables",
            selectInput("logistic_dep", "Dependent", choices = NULL),
            selectizeInput("logistic_indep", "Independent(s)", choices = NULL, multiple = TRUE),
            uiOutput("logistic_ref_ui"),
            actionButton("run_logistic", "OK", class = "btn-primary", icon = icon("check")),
            hr(),
            h4("Method"),
            selectInput("logistic_method", "Method",
                        choices = c("Enter", "Forward", "Backward", "Stepwise")),
            h4("Options"),
            checkboxGroupInput("logistic_options", "Display",
                               choices = c("CI for exp(B)", "Hosmer-Lemeshow goodness-of-fit", 
                                           "Iteration history", "Correlation of estimates")),
            hr(),
            h4("Export Options"),
            downloadButton("export_logistic_word", "Word", class = "export-btn"),
            downloadButton("export_logistic_excel", "Excel", class = "export-btn")
          ),
          box(
            width = 8, status = "info", title = "Logistic Regression",
            tabBox(
              width = 12,
              tabPanel("Model Summary", div(class = "result-output", withSpinner(verbatimTextOutput("logistic_summary")))),
              tabPanel("Coefficients", div(class = "result-output", withSpinner(verbatimTextOutput("logistic_coef")))),
              tabPanel("Classification", div(class = "result-output", withSpinner(verbatimTextOutput("logistic_classification")))),
              tabPanel("H-L Test", div(class = "result-output", withSpinner(verbatimTextOutput("logistic_hl")))),
              tabPanel("Plots", withSpinner(plotlyOutput("logistic_plot", height = "400px")))
            )
          )
        )
      ),
      
      # Ordinal Regression Tab
      tabItem(
        tabName = "ordinal_reg",
        h2("Ordinal Regression", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Variables",
            selectInput("ordinal_dep", "Dependent", choices = NULL),
            selectizeInput("ordinal_indep", "Independent(s)", choices = NULL, multiple = TRUE),
            actionButton("run_ordinal", "OK", class = "btn-primary", icon = icon("check")),
            hr(),
            h4("Export Options"),
            downloadButton("export_ordinal_word", "Word", class = "export-btn"),
            downloadButton("export_ordinal_excel", "Excel", class = "export-btn")
          ),
          box(
            width = 8, status = "info", title = "Ordinal Regression",
            div(class = "result-output",
                withSpinner(verbatimTextOutput("ordinal_output"))
            )
          )
        )
      ),
      
      # Multinomial Regression Tab
      tabItem(
        tabName = "multinomial_reg",
        h2("Multinomial Regression", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Variables",
            selectInput("multinomial_dep", "Dependent", choices = NULL),
            selectizeInput("multinomial_indep", "Independent(s)", choices = NULL, multiple = TRUE),
            uiOutput("multinomial_ref_ui"),
            actionButton("run_multinomial", "OK", class = "btn-primary", icon = icon("check")),
            hr(),
            h4("Export Options"),
            downloadButton("export_multinomial_word", "Word", class = "export-btn"),
            downloadButton("export_multinomial_excel", "Excel", class = "export-btn")
          ),
          box(
            width = 8, status = "info", title = "Multinomial Regression",
            div(class = "result-output",
                withSpinner(verbatimTextOutput("multinomial_output"))
            )
          )
        )
      ),
      
      # Poisson Regression Tab
      tabItem(
        tabName = "poisson_reg",
        h2("Poisson Regression", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Variables",
            selectInput("poisson_dep", "Dependent", choices = NULL),
            selectizeInput("poisson_indep", "Independent(s)", choices = NULL, multiple = TRUE),
            actionButton("run_poisson", "OK", class = "btn-primary", icon = icon("check")),
            hr(),
            h4("Export Options"),
            downloadButton("export_poisson_word", "Word", class = "export-btn"),
            downloadButton("export_poisson_excel", "Excel", class = "export-btn")
          ),
          box(
            width = 8, status = "info", title = "Poisson Regression",
            div(class = "result-output",
                withSpinner(verbatimTextOutput("poisson_output"))
            )
          )
        )
      ),
      
      # Factor Analysis Tab
      tabItem(
        tabName = "factor_analysis",
        h2("Exploratory Factor Analysis", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Variables",
            selectizeInput("factor_vars", "Variables", choices = NULL, multiple = TRUE),
            numericInput("num_factors", "Number of Factors", value = 2, min = 1, max = 20),
            selectInput("rotation_method", "Rotation Method",
                        choices = c("none", "varimax", "quartimax", "promax", "oblimin", "simplimax", "cluster")),
            actionButton("run_factor", "OK", class = "btn-primary", icon = icon("check")),
            hr(),
            h4("Options"),
            checkboxGroupInput("factor_options", "Display",
                               choices = c("Eigenvalues", "Factor loadings", "Communalities", "Score coefficients")),
            hr(),
            h4("Export Options"),
            downloadButton("export_factor_word", "Word", class = "export-btn"),
            downloadButton("export_factor_excel", "Excel", class = "export-btn")
          ),
          box(
            width = 8, status = "info", title = "Factor Analysis",
            tabBox(
              width = 12,
              tabPanel("Summary", div(class = "result-output", withSpinner(verbatimTextOutput("factor_summary")))),
              tabPanel("Loadings", div(class = "result-output", withSpinner(verbatimTextOutput("factor_loadings")))),
              tabPanel("Scree Plot", withSpinner(plotlyOutput("factor_scree", height = "400px"))),
              tabPanel("Factor Scores", div(class = "result-output", withSpinner(verbatimTextOutput("factor_scores"))))
            )
          )
        )
      ),
      
      # CFA Tab
      tabItem(
        tabName = "cfa",
        h2("Confirmatory Factor Analysis", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Model Specification",
            selectizeInput("cfa_vars", "Variables", choices = NULL, multiple = TRUE),
            textAreaInput("cfa_model", "Model Syntax", 
                          placeholder = "f1 =~ var1 + var2 + var3\nf2 =~ var4 + var5 + var6",
                          rows = 5),
            actionButton("run_cfa", "OK", class = "btn-primary", icon = icon("check")),
            hr(),
            h4("Options"),
            checkboxGroupInput("cfa_options", "Display",
                               choices = c("Fit measures", "Parameter estimates", "Standardized results", "Modification indices")),
            hr(),
            h4("Export Options"),
            downloadButton("export_cfa_word", "Word", class = "export-btn"),
            downloadButton("export_cfa_excel", "Excel", class = "export-btn")
          ),
          box(
            width = 8, status = "info", title = "Confirmatory Factor Analysis",
            tabBox(
              width = 12,
              tabPanel("Summary", div(class = "result-output", withSpinner(verbatimTextOutput("cfa_summary")))),
              tabPanel("Fit Measures", div(class = "result-output", withSpinner(verbatimTextOutput("cfa_fit")))),
              tabPanel("Parameter Estimates", div(class = "result-output", withSpinner(verbatimTextOutput("cfa_params"))))
            )
          )
        )
      ),
      
      # Reliability Analysis Tab
      tabItem(
        tabName = "reliability",
        h2("Reliability Analysis", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Variables",
            selectizeInput("reliability_vars", "Items", choices = NULL, multiple = TRUE),
            actionButton("run_reliability", "OK", class = "btn-primary", icon = icon("check")),
            hr(),
            h4("Statistics"),
            checkboxGroupInput("reliability_stats", "Display",
                               choices = c("Cronbach's alpha", "McDonald's omega", "Item statistics", "Scale statistics")),
            hr(),
            h4("Export Options"),
            downloadButton("export_reliability_word", "Word", class = "export-btn"),
            downloadButton("export_reliability_excel", "Excel", class = "export-btn")
          ),
          box(
            width = 8, status = "info", title = "Reliability Analysis",
            div(class = "result-output",
                withSpinner(verbatimTextOutput("reliability_output"))
            )
          )
        )
      ),
      
      # Chi-Square Tab
      tabItem(
        tabName = "chi_square",
        h2("Chi-Square Test", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Variables",
            selectInput("chi_row", "Row Variable", choices = NULL),
            selectInput("chi_col", "Column Variable", choices = NULL),
            actionButton("run_chi", "OK", class = "btn-primary", icon = icon("check")),
            hr(),
            h4("Statistics"),
            checkboxGroupInput("chi_stats", "Display",
                               choices = c("Chi-square", "Phi and Cramer's V", "Contingency coefficient", 
                                           "Lambda", "Uncertainty coefficient")),
            hr(),
            h4("Export Options"),
            downloadButton("export_chi_word", "Word", class = "export-btn"),
            downloadButton("export_chi_excel", "Excel", class = "export-btn")
          ),
          box(
            width = 8, status = "info", title = "Chi-Square Test",
            tabBox(
              width = 12,
              tabPanel("Crosstabulation", div(class = "result-output", withSpinner(verbatimTextOutput("chi_crosstab")))),
              tabPanel("Test Statistics", div(class = "result-output", withSpinner(verbatimTextOutput("chi_test")))),
              tabPanel("Plots", withSpinner(plotlyOutput("chi_plot", height = "400px")))
            )
          )
        )
      ),
      
      # Mann-Whitney U Test Tab
      tabItem(
        tabName = "mann_whitney",
        h2("Mann-Whitney U Test", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Variables",
            selectizeInput("mw_vars", "Test Variable(s)", choices = NULL, multiple = TRUE),
            selectInput("mw_group_var", "Grouping Variable", choices = NULL),
            textInput("mw_group_def", "Define Groups", placeholder = "e.g., 1,2 or A,B"),
            actionButton("run_mw", "OK", class = "btn-primary", icon = icon("check")),
            hr(),
            h4("Export Options"),
            downloadButton("export_mw_word", "Word", class = "export-btn"),
            downloadButton("export_mw_excel", "Excel", class = "export-btn")
          ),
          box(
            width = 8, status = "info", title = "Mann-Whitney U Test",
            div(class = "result-output",
                withSpinner(verbatimTextOutput("mw_output"))
            )
          )
        )
      ),
      
      # Kruskal-Wallis Test Tab
      tabItem(
        tabName = "kruskal_wallis",
        h2("Kruskal-Wallis Test", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Variables",
            selectizeInput("kw_vars", "Test Variable(s)", choices = NULL, multiple = TRUE),
            selectInput("kw_group_var", "Grouping Variable", choices = NULL),
            actionButton("run_kw", "OK", class = "btn-primary", icon = icon("check")),
            hr(),
            h4("Export Options"),
            downloadButton("export_kw_word", "Word", class = "export-btn"),
            downloadButton("export_kw_excel", "Excel", class = "export-btn")
          ),
          box(
            width = 8, status = "info", title = "Kruskal-Wallis Test",
            div(class = "result-output",
                withSpinner(verbatimTextOutput("kw_output"))
            )
          )
        )
      ),
      
      # Wilcoxon Signed-Rank Test Tab
      tabItem(
        tabName = "wilcoxon",
        h2("Wilcoxon Signed-Rank Test", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Variables",
            selectizeInput("wilcoxon_vars", "Paired Variables", choices = NULL, multiple = TRUE),
            actionButton("run_wilcoxon", "OK", class = "btn-primary", icon = icon("check")),
            hr(),
            h4("Export Options"),
            downloadButton("export_wilcoxon_word", "Word", class = "export-btn"),
            downloadButton("export_wilcoxon_excel", "Excel", class = "export-btn")
          ),
          box(
            width = 8, status = "info", title = "Wilcoxon Signed-Rank Test",
            div(class = "result-output",
                withSpinner(verbatimTextOutput("wilcoxon_output"))
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
                                    "Boxplot", "Density", "Violin", "Q-Q Plot")),
            hr(),
            h4("Customization"),
            textInput("chart_title", "Title", value = ""),
            numericInput("title_size", "Title Size", value = 14, min = 8, max = 24, step = 1),
            numericInput("axis_size", "Axis Text Size", value = 12, min = 8, max = 20, step = 1),
            selectInput("color_palette", "Color Palette", 
                        choices = c("SPSS Blue" = "spss", "Set1", "Set2", "Set3", "Pastel1", "Pastel2", "Dark2", "Accent")),
            numericInput("chart_width", "Width (px)", value = 800, min = 400, max = 2000, step = 50),
            numericInput("chart_height", "Height (px)", value = 500, min = 300, max = 1500, step = 50),
            hr(),
            actionButton("generate_chart", "Generate Chart", class = "btn-primary", icon = icon("check")),
            downloadButton("download_chart", "Download PNG", class = "export-btn")
          ),
          box(
            width = 9, status = "info", title = "Chart Preview",
            withSpinner(plotOutput("chart_output", height = "500px"))
          )
        )
      ),
      
      # Help Tab
      tabItem(
        tabName = "help",
        h2("Statistical Test Selection Guide", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 12, status = "info", title = "How to Choose Appropriate Statistical Tests",
            div(class = "analysis-panel",
                h3("Variable Types and Appropriate Tests"),
                p("Selecting the right statistical test depends on the type of variables you have and the research question you want to answer."),
                
                h4("1. Categorical vs Categorical Variables"),
                p("When both variables are categorical:"),
                tags$ul(
                  tags$li("Chi-Square Test: Tests for association between two categorical variables"),
                  tags$li("Fisher's Exact Test: Alternative to chi-square when sample sizes are small"),
                  tags$li("Log-Linear Analysis: For analyzing multi-way contingency tables")
                ),
                
                h4("2. Continuous vs Categorical Variables"),
                p("When comparing a continuous variable across categories:"),
                tags$ul(
                  tags$li("T-Test: Compare means between two groups (Independent or Paired)"),
                  tags$li("ANOVA: Compare means among three or more groups"),
                  tags$li("Mann-Whitney U Test: Nonparametric alternative to independent t-test"),
                  tags$li("Kruskal-Wallis Test: Nonparametric alternative to one-way ANOVA"),
                  tags$li("Wilcoxon Signed-Rank Test: Nonparametric alternative to paired t-test")
                ),
                
                h4("3. Continuous vs Continuous Variables"),
                p("When examining relationships between continuous variables:"),
                tags$ul(
                  tags$li("Pearson Correlation: Measures linear relationship between two continuous variables"),
                  tags$li("Spearman Correlation: Nonparametric measure of monotonic relationship"),
                  tags$li("Simple Linear Regression: Predicts one continuous variable from another"),
                  tags$li("Multiple Regression: Predicts one continuous variable from multiple predictors")
                ),
                
                h4("4. Predicting Categorical Outcomes"),
                p("When the outcome variable is categorical:"),
                tags$ul(
                  tags$li("Logistic Regression: For binary outcomes (2 categories)"),
                  tags$li("Multinomial Regression: For nominal outcomes (3+ categories)"),
                  tags$li("Ordinal Regression: For ordinal outcomes (ordered categories)")
                ),
                
                h4("5. Count Data"),
                p("When working with count data:"),
                tags$ul(
                  tags$li("Poisson Regression: For count data where variance equals mean"),
                  tags$li("Negative Binomial Regression: For overdispersed count data")
                ),
                
                h4("6. Dimension Reduction"),
                p("When you want to reduce the number of variables:"),
                tags$ul(
                  tags$li("Factor Analysis: Identifies underlying factors that explain patterns in variables"),
                  tags$li("Principal Component Analysis: Transforms variables into uncorrelated components")
                ),
                
                h4("7. Reliability Analysis"),
                p("When assessing the reliability of measurements:"),
                tags$ul(
                  tags$li("Cronbach's Alpha: Measures internal consistency of a scale"),
                  tags$li("McDonald's Omega: Alternative reliability coefficient that may be more appropriate")
                ),
                
                h4("8. Model Assumptions"),
                p("Always check assumptions before interpreting results:"),
                tags$ul(
                  tags$li("Normality: Shapiro-Wilk test, Q-Q plots"),
                  tags$li("Homogeneity of Variance: Levene's test, Bartlett's test"),
                  tags$li("Linearity: Scatterplots, residual plots"),
                  tags$li("Independence: Durbin-Watson test (for regression)")
                ),
                
                p("For more detailed guidance, consult statistical textbooks or consult with a statistician.")
            )
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
  data_modified <- reactiveVal(FALSE)
  
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
    data_modified(FALSE)
    
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
    update_all_select_inputs(df, types, session)
    
    showNotification("Example data loaded successfully!", type = "message")
  })
  
  # Data import and processing
  observeEvent(input$import_btn, {
    req(input$file)
    
    ext <- tools::file_ext(input$file$name)
    
    tryCatch({
      df <- switch(ext,
                   csv = read.csv(input$file$datapath, 
                                  header = input$header, 
                                  sep = input$sep,
                                  stringsAsFactors = TRUE),
                   xlsx = read_excel(input$file$datapath),
                   xls = read_excel(input$file$datapath),
                   sav = read_sav(input$file$datapath),
                   dta = read_dta(input$file$datapath),
                   sas7bdat = read_sas(input$file$datapath),
                   rdata = {
                     env <- new.env()
                     load(input$file$datapath, envir = env)
                     get(ls(env)[1], envir = env)
                   },
                   rda = {
                     env <- new.env()
                     load(input$file$datapath, envir = env)
                     get(ls(env)[1], envir = env)
                   },
                   stop("Invalid file type")
      )
      
      # Convert to data frame if necessary
      if (!is.data.frame(df)) {
        df <- as.data.frame(df)
      }
      
      # Convert character columns to factors
      char_cols <- sapply(df, is.character)
      if (any(char_cols)) {
        df[char_cols] <- lapply(df[char_cols], as.factor)
      }
      
      data(df)
      data_modified(FALSE)
      
      # Determine variable types
      types <- sapply(df, function(x) {
        if (is.numeric(x)) "Numeric"
        else if (is.factor(x)) "Categorical"
        else "Other"
      })
      
      var_types(data.frame(
        Variable = names(df),
        Type = types,
        stringsAsFactors = FALSE
      ))
      
      # Update all select inputs
      update_all_select_inputs(df, types, session)
      
      showNotification("Data imported successfully!", type = "message")
      
    }, error = function(e) {
      showNotification(paste("Error:", e$message), type = "error")
    })
  })
  
  # Function to update all select inputs
  update_all_select_inputs <- function(df, types, session) {
    updateSelectizeInput(session, "freq_vars", choices = names(df))
    updateSelectizeInput(session, "desc_vars", choices = names(df))
    updateSelectizeInput(session, "ttest_vars", choices = names(df))
    updateSelectInput(session, "ttest_group_var", choices = names(df)[types == "Categorical"])
    updateSelectizeInput(session, "onesample_vars", choices = names(df))
    updateSelectizeInput(session, "paired_vars", choices = names(df))
    updateSelectizeInput(session, "anova_dv", choices = names(df)[types == "Numeric"])
    updateSelectInput(session, "anova_factor", choices = names(df)[types == "Categorical"])
    updateSelectizeInput(session, "cor_vars", choices = names(df)[types == "Numeric"])
    updateSelectInput(session, "reg_dep", choices = names(df)[types == "Numeric"])
    updateSelectizeInput(session, "reg_indep", choices = names(df))
    updateSelectInput(session, "logistic_dep", choices = names(df)[types == "Categorical"])
    updateSelectizeInput(session, "logistic_indep", choices = names(df))
    updateSelectInput(session, "ordinal_dep", choices = names(df)[types == "Categorical"])
    updateSelectizeInput(session, "ordinal_indep", choices = names(df))
    updateSelectInput(session, "multinomial_dep", choices = names(df)[types == "Categorical"])
    updateSelectizeInput(session, "multinomial_indep", choices = names(df))
    updateSelectInput(session, "poisson_dep", choices = names(df)[types == "Numeric"])
    updateSelectizeInput(session, "poisson_indep", choices = names(df))
    updateSelectInput(session, "chi_row", choices = names(df)[types == "Categorical"])
    updateSelectInput(session, "chi_col", choices = names(df)[types == "Categorical"])
    updateSelectizeInput(session, "factor_vars", choices = names(df)[types == "Numeric"])
    updateSelectizeInput(session, "chart_vars", choices = names(df))
    updateSelectInput(session, "chart_group", choices = c("None", names(df)[types == "Categorical"]))
    updateSelectInput(session, "define_var", choices = names(df))
    updateSelectizeInput(session, "sort_vars", choices = names(df))
    updateSelectInput(session, "weight_var", choices = names(df)[types == "Numeric"])
    updateSelectInput(session, "recode_var", choices = names(df))
    updateSelectInput(session, "bin_var", choices = names(df)[types == "Numeric"])
    updateSelectInput(session, "log_var", choices = names(df)[types == "Numeric"])
    updateSelectInput(session, "standardize_var", choices = names(df)[types == "Numeric"])
    updateSelectizeInput(session, "explore_vars", choices = names(df)[types == "Numeric"])
    updateSelectInput(session, "explore_factor", choices = c("None", names(df)[types == "Categorical"]))
    updateSelectInput(session, "twoway_dv", choices = names(df)[types == "Numeric"])
    updateSelectInput(session, "twoway_factor1", choices = names(df)[types == "Categorical"])
    updateSelectInput(session, "twoway_factor2", choices = names(df)[types == "Categorical"])
    updateSelectizeInput(session, "cfa_vars", choices = names(df))
    updateSelectizeInput(session, "reliability_vars", choices = names(df))
    updateSelectizeInput(session, "mw_vars", choices = names(df)[types == "Numeric"])
    updateSelectInput(session, "mw_group_var", choices = names(df)[types == "Categorical"])
    updateSelectizeInput(session, "kw_vars", choices = names(df)[types == "Numeric"])
    updateSelectInput(session, "kw_group_var", choices = names(df)[types == "Categorical"])
    updateSelectizeInput(session, "wilcoxon_vars", choices = names(df)[types == "Numeric"])
  }
  
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
  
  # Save data preview
  output$save_preview <- renderDT({
    req(data())
    datatable(data(), options = list(scrollX = TRUE, pageLength = 5))
  })
  
  # Save data functionality
  output$save_btn <- downloadHandler(
    filename = function() {
      paste(input$save_filename, switch(input$save_format,
                                        "CSV" = ".csv",
                                        "Excel" = ".xlsx",
                                        "SPSS" = ".sav",
                                        "Stata" = ".dta",
                                        "RData" = ".RData"), sep = "")
    },
    content = function(file) {
      req(data())
      df <- data()
      
      switch(input$save_format,
             "CSV" = write.csv(df, file, row.names = FALSE),
             "Excel" = write.xlsx(df, file),
             "SPSS" = write_sav(df, file),
             "Stata" = write_dta(df, file),
             "RData" = save(df, file = file))
      
      showNotification("Data saved successfully!", type = "message")
    }
  )
  
  # Define Variable Properties functionality
  output$var_props_output <- renderPrint({
    req(data(), input$define_var)
    
    df <- data()
    var <- df[[input$define_var]]
    
    cat("Variable: ", input$define_var, "\n")
    cat("Current Type: ", class(var), "\n")
    cat("Number of Missing Values: ", sum(is.na(var)), "\n")
    cat("Number of Valid Values: ", sum(!is.na(var)), "\n")
    
    if (is.numeric(var)) {
      cat("Minimum: ", min(var, na.rm = TRUE), "\n")
      cat("Maximum: ", max(var, na.rm = TRUE), "\n")
      cat("Mean: ", mean(var, na.rm = TRUE), "\n")
      cat("Standard Deviation: ", sd(var, na.rm = TRUE), "\n")
    } else if (is.factor(var) || is.character(var)) {
      cat("Levels/Categories: ", paste(unique(var), collapse = ", "), "\n")
    }
  })
  
  observeEvent(input$apply_var_props, {
    req(data(), input$define_var)
    
    df <- data()
    
    # Apply variable label
    if (input$var_label != "") {
      # In R, we can use attributes or comment to store variable labels
      attr(df[[input$define_var]], "label") <- input$var_label
    }
    
    # Apply variable type
    if (input$var_type == "Categorical" && is.numeric(df[[input$define_var]])) {
      df[[input$define_var]] <- as.factor(df[[input$define_var]])
    } else if (input$var_type == "Numeric" && (is.character(df[[input$define_var]]) || is.factor(df[[input$define_var]]))) {
      df[[input$define_var]] <- as.numeric(as.character(df[[input$define_var]]))
    }
    
    # Apply value labels
    if (input$var_type == "Categorical" && input$value_labels != "") {
      labels <- strsplit(input$value_labels, ",")[[1]]
      for (label in labels) {
        parts <- strsplit(label, "=")[[1]]
        if (length(parts) == 2) {
          level <- trimws(parts[1])
          label_text <- trimws(parts[2])
          levels(df[[input$define_var]])[levels(df[[input$define_var]]) == level] <- label_text
        }
      }
    }
    
    data(df)
    data_modified(TRUE)
    
    # Update variable types
    types <- sapply(df, function(x) {
      if (is.numeric(x)) "Numeric"
      else if (is.factor(x)) "Categorical"
      else "Other"
    })
    
    var_types(data.frame(
      Variable = names(df),
      Type = types,
      stringsAsFactors = FALSE
    ))
    
    showNotification("Variable properties applied successfully!", type = "message")
  })
  
  # Sort Cases functionality
  observeEvent(input$sort_btn, {
    req(data(), input$sort_vars)
    
    df <- data()
    
    # Create sorting formula
    sort_expr <- paste0("df[order(", paste0("df$", input$sort_vars, 
                                            ifelse(input$sort_order == "Descending", ", decreasing = TRUE)", ")"), 
                                            collapse = ", "), "]")
    
    sorted_df <- eval(parse(text = sort_expr))
    data(sorted_df)
    data_modified(TRUE)
    
    output$sorted_preview <- renderDT({
      datatable(sorted_df, options = list(scrollX = TRUE, pageLength = 5))
    })
    
    showNotification("Cases sorted successfully!", type = "message")
  })
  
  # Select Cases functionality
  observeEvent(input$select_btn, {
    req(data())
    
    df <- data()
    
    if (input$select_method == "If condition is satisfied" && input$select_condition != "") {
      # Filter based on condition
      condition <- input$select_condition
      filtered_df <- df %>% filter(eval(parse(text = condition)))
      data(filtered_df)
      data_modified(TRUE)
      
      output$selected_preview <- renderDT({
        datatable(filtered_df, options = list(scrollX = TRUE, pageLength = 5))
      })
      
      showNotification(paste("Selected", nrow(filtered_df), "cases based on condition"), type = "message")
    } else if (input$select_method == "Random sample") {
      # Take random sample
      sample_size <- round(nrow(df) * input$sample_percent / 100)
      sampled_df <- df %>% sample_n(sample_size)
      data(sampled_df)
      data_modified(TRUE)
      
      output$selected_preview <- renderDT({
        datatable(sampled_df, options = list(scrollX = TRUE, pageLength = 5))
      })
      
      showNotification(paste("Selected", sample_size, "random cases"), type = "message")
    } else {
      # Reset to all cases
      data(df)
      data_modified(FALSE)
      
      output$selected_preview <- renderDT({
        datatable(df, options = list(scrollX = TRUE, pageLength = 5))
      })
      
      showNotification("All cases selected", type = "message")
    }
  })
  
  # Weight Cases functionality
  observeEvent(input$weight_btn, {
    req(data())
    
    df <- data()
    
    if (input$weight_method == "Weight cases by" && !is.null(input$weight_var)) {
      # Apply weighting
      weights <- df[[input$weight_var]]
      
      # Store original data and apply weights
      attr(df, "weights") <- weights
      
      output$weight_output <- renderPrint({
        cat("Weighting applied using variable:", input$weight_var, "\n")
        cat("Summary of weights:\n")
        print(summary(weights))
      })
      
      showNotification("Weighting applied successfully!", type = "message")
    } else {
      # Remove weighting
      attr(df, "weights") <- NULL
      
      output$weight_output <- renderPrint({
        cat("No weighting applied\n")
      })
      
      showNotification("Weighting removed", type = "message")
    }
    
    data(df)
  })
  
  # Transform functionality
  observeEvent(input$apply_transform, {
    req(data())
    
    df <- data()
    
    tryCatch({
      if (input$transform_type == "Recode") {
        req(input$recode_var, input$recode_rules)
        
        # Parse recode rules
        rules <- strsplit(input$recode_rules, ",")[[1]]
        recode_list <- list()
        
        for (rule in rules) {
          parts <- strsplit(rule, "=")[[1]]
          if (length(parts) == 2) {
            old_val <- trimws(parts[1])
            new_val <- trimws(parts[2])
            
            # Convert to appropriate type
            if (is.numeric(df[[input$recode_var]])) {
              old_val <- as.numeric(old_val)
              new_val <- as.numeric(new_val)
            }
            
            recode_list[[as.character(old_val)]] <- new_val
          }
        }
        
        # Apply recoding
        if (input$recode_into == "Same variable") {
          df[[input$recode_var]] <- recode(df[[input$recode_var]], !!!recode_list)
        } else {
          new_var_name <- ifelse(input$new_recode_var != "", input$new_recode_var, 
                                 paste0(input$recode_var, "_recoded"))
          df[[new_var_name]] <- recode(df[[input$recode_var]], !!!recode_list)
        }
        
      } else if (input$transform_type == "Compute") {
        req(input$compute_var, input$compute_expr)
        
        # Parse and evaluate expression
        expr <- parse(text = input$compute_expr)
        df[[input$compute_var]] <- eval(expr, envir = df)
        
      } else if (input$transform_type == "Binning") {
        req(input$bin_var, input$binned_var)
        
        # Create bins
        df[[input$binned_var]] <- cut(df[[input$bin_var]], breaks = input$num_bins, 
                                      labels = paste0("Bin", 1:input$num_bins))
        
      } else if (input$transform_type == "Log Transform") {
        req(input$log_var, input$log_var_name)
        
        # Apply log transformation
        if (input$log_type == "Natural Log") {
          df[[input$log_var_name]] <- log(df[[input$log_var]])
        } else {
          df[[input$log_var_name]] <- log10(df[[input$log_var]])
        }
        
      } else if (input$transform_type == "Standardize") {
        req(input$standardize_var, input$z_var_name)
        
        # Standardize variable
        df[[input$z_var_name]] <- scale(df[[input$standardize_var]])
      }
      
      data(df)
      data_modified(TRUE)
      
      # Update variable types
      types <- sapply(df, function(x) {
        if (is.numeric(x)) "Numeric"
        else if (is.factor(x)) "Categorical"
        else "Other"
      })
      
      var_types(data.frame(
        Variable = names(df),
        Type = types,
        stringsAsFactors = FALSE
      ))
      
      # Update all select inputs
      update_all_select_inputs(df, types, session)
      
      output$transform_preview <- renderDT({
        datatable(df, options = list(scrollX = TRUE, pageLength = 5))
      })
      
      showNotification("Transformation applied successfully!", type = "message")
      
    }, error = function(e) {
      showNotification(paste("Error:", e$message), type = "error")
    })
  })
  
  # Reference category UI for logistic regression
  output$logistic_ref_ui <- renderUI({
    req(data(), input$logistic_dep)
    df <- data()
    dep_var <- df[[input$logistic_dep]]
    if (is.factor(dep_var)) {
      choices <- levels(dep_var)
    } else {
      choices <- unique(dep_var)
    }
    selectInput("logistic_ref", "Reference Category", choices = choices)
  })
  
  # Reference category UI for multinomial regression
  output$multinomial_ref_ui <- renderUI({
    req(data(), input$multinomial_dep)
    df <- data()
    dep_var <- df[[input$multinomial_dep]]
    if (is.factor(dep_var)) {
      choices <- levels(dep_var)
    } else {
      choices <- unique(dep_var)
    }
    selectInput("multinomial_ref", "Reference Category", choices = choices)
  })
  
  # Descriptive statistics
  output$desc_output <- renderPrint({
    req(data(), input$desc_vars)
    
    df <- data()
    vars <- df[, input$desc_vars, drop = FALSE]
    
    if (ncol(vars) == 0) return("Please select at least one variable")
    
    # Calculate requested statistics
    results <- list()
    
    if ("Valid N" %in% input$desc_stats) {
      valid_n <- sapply(vars, function(x) sum(!is.na(x)))
      results$N <- valid_n
    }
    
    if ("Missing N" %in% input$desc_stats) {
      missing_n <- sapply(vars, function(x) sum(is.na(x)))
      results$`Missing N` <- missing_n
    }
    
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
  
  # Export descriptive statistics to Word
  output$export_descriptives_word <- downloadHandler(
    filename = function() {
      paste("descriptive_stats_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(data(), input$desc_vars)
      
      df <- data()
      vars <- df[, input$desc_vars, drop = FALSE]
      
      # Calculate requested statistics
      results <- list()
      
      if ("Valid N" %in% input$desc_stats) {
        valid_n <- sapply(vars, function(x) sum(!is.na(x)))
        results$N <- valid_n
      }
      
      if ("Missing N" %in% input$desc_stats) {
        missing_n <- sapply(vars, function(x) sum(is.na(x)))
        results$`Missing N` <- missing_n
      }
      
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
      
      # Create APA-style table
      result_df <- as.data.frame(t(do.call(rbind, results)))
      result_df <- cbind(Variable = rownames(result_df), result_df)
      rownames(result_df) <- NULL
      
      # Create flextable
      ft <- flextable(result_df)
      ft <- set_caption(ft, caption = "Descriptive Statistics")
      ft <- theme_apa(ft)
      ft <- autofit(ft)
      
      # Save to Word
      doc <- read_docx()
      doc <- body_add_flextable(doc, value = ft)
      print(doc, target = file)
    }
  )
  
  # Export descriptive statistics to Excel
  output$export_descriptives_excel <- downloadHandler(
    filename = function() {
      paste("descriptive_stats_", Sys.Date(), ".xlsx", sep = "")
    },
    content = function(file) {
      req(data(), input$desc_vars)
      
      df <- data()
      vars <- df[, input$desc_vars, drop = FALSE]
      
      # Calculate requested statistics
      results <- list()
      
      if ("Valid N" %in% input$desc_stats) {
        valid_n <- sapply(vars, function(x) sum(!is.na(x)))
        results$N <- valid_n
      }
      
      if ("Missing N" %in% input$desc_stats) {
        missing_n <- sapply(vars, function(x) sum(is.na(x)))
        results$`Missing N` <- missing_n
      }
      
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
      
      # Create data frame
      result_df <- as.data.frame(t(do.call(rbind, results)))
      result_df <- cbind(Variable = rownames(result_df), result_df)
      rownames(result_df) <- NULL
      
      # Write to Excel
      write.xlsx(result_df, file)
    }
  )
  
  # Frequency tables
  output$freq_output <- renderPrint({
    req(data(), input$freq_vars)
    
    if (length(input$freq_vars) == 0) return("Please select at least one variable")
    
    cat("Frequencies\n")
    cat("===========\n\n")
    
    for (var_name in input$freq_vars) {
      var <- data()[[var_name]]
      
      freq_table <- table(var, useNA = "ifany")
      prop_table <- prop.table(freq_table) * 100
      valid_prop <- prop.table(table(var)) * 100
      cum_table <- cumsum(valid_prop)
      
      cat("Variable: ", var_name, "\n\n")
      
      result_df <- data.frame(
        Frequency = as.numeric(freq_table),
        Percent = round(as.numeric(prop_table), 1),
        ValidPercent = round(as.numeric(valid_prop), 1),
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
        if ("Mode" %in% input$freq_stats) {
          mode_val <- names(sort(table(var), decreasing = TRUE))[1]
          stats_list$Mode <- mode_val
        }
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
  
  # Export frequencies to Word
  output$export_frequencies_word <- downloadHandler(
    filename = function() {
      paste("frequencies_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(data(), input$freq_vars)
      
      if (length(input$freq_vars) == 0) return()
      
      doc <- read_docx()
      
      for (var_name in input$freq_vars) {
        var <- data()[[var_name]]
        
        freq_table <- table(var, useNA = "ifany")
        prop_table <- prop.table(freq_table) * 100
        valid_prop <- prop.table(table(var)) * 100
        cum_table <- cumsum(valid_prop)
        
        result_df <- data.frame(
          Frequency = as.numeric(freq_table),
          Percent = round(as.numeric(prop_table), 1),
          ValidPercent = round(as.numeric(valid_prop), 1),
          CumulativePercent = round(as.numeric(cum_table), 1)
        )
        
        # Add value labels if available
        if (is.factor(var)) {
          result_df <- cbind(Value = levels(var), result_df)
        } else {
          result_df <- cbind(Value = names(freq_table), result_df)
        }
        
        # Create flextable
        ft <- flextable(result_df)
        ft <- set_caption(ft, caption = paste("Frequencies for", var_name))
        ft <- theme_apa(ft)
        ft <- autofit(ft)
        
        doc <- body_add_flextable(doc, value = ft)
        doc <- body_add_par(doc, "")
      }
      
      print(doc, target = file)
    }
  )
  
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
    alternative <- ifelse(input$ttest_tails == "One-tailed", "greater", "two.sided")
    
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
        conf.level = conf_level,
        alternative = alternative
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
  
  # Export t-test results to Word
  output$export_ttest_word <- downloadHandler(
    filename = function() {
      paste("ttest_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(data(), input$ttest_vars, input$ttest_group_var, input$group_def)
      
      # Parse group definition
      groups <- trimws(unlist(strsplit(input$group_def, ",")))
      if (length(groups) != 2) return()
      
      # Filter data based on groups
      group_var <- data()[[input$ttest_group_var]]
      if (!all(groups %in% unique(group_var))) return()
      
      # Create subset for each group
      group1_data <- data()[group_var == groups[1], input$ttest_vars, drop = FALSE]
      group2_data <- data()[group_var == groups[2], input$ttest_vars, drop = FALSE]
      
      # Perform t-test for each variable
      conf_level <- input$ttest_conf / 100
      alternative <- ifelse(input$ttest_tails == "One-tailed", "greater", "two.sided")
      
      results_list <- list()
      
      for (var in input$ttest_vars) {
        if (!is.numeric(data()[[var]])) next
        
        test_result <- t.test(
          group1_data[[var]], 
          group2_data[[var]], 
          conf.level = conf_level,
          alternative = alternative
        )
        
        results_list[[var]] <- data.frame(
          Variable = var,
          t.value = round(test_result$statistic, 3),
          df = round(test_result$parameter, 1),
          p.value = format.pval(test_result$p.value, digits = 3),
          Mean.Difference = round(diff(test_result$estimate), 3),
          CI.Lower = round(test_result$conf.int[1], 3),
          CI.Upper = round(test_result$conf.int[2], 3)
        )
      }
      
      if (length(results_list) == 0) return()
      
      result_df <- do.call(rbind, results_list)
      
      # Create flextable
      ft <- flextable(result_df)
      ft <- set_caption(ft, caption = "Independent Samples T-Test")
      ft <- theme_apa(ft)
      ft <- autofit(ft)
      
      # Save to Word
      doc <- read_docx()
      doc <- body_add_flextable(doc, value = ft)
      print(doc, target = file)
    }
  )
  
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
  
  # One Sample T-Test output
  output$onesample_output <- renderPrint({
    req(data(), input$onesample_vars)
    
    if (length(input$onesample_vars) == 0) return("Please select at least one variable")
    
    conf_level <- input$onesample_conf / 100
    alternative <- ifelse(input$onesample_tails == "One-tailed", "greater", "two.sided")
    
    cat("One Sample T-Test\n")
    cat("=================\n\n")
    cat("Test Value =", input$test_value, "\n\n")
    
    for (var in input$onesample_vars) {
      cat("Variable: ", var, "\n")
      cat("------------------------\n")
      
      # Check if variable is numeric
      if (!is.numeric(data()[[var]])) {
        cat("Skipping non-numeric variable:", var, "\n\n")
        next
      }
      
      # Perform one-sample t-test
      test_result <- t.test(
        data()[[var]], 
        mu = input$test_value,
        conf.level = conf_level,
        alternative = alternative
      )
      
      # Format output
      cat("t-value:", round(test_result$statistic, 3), "\n")
      cat("df:", round(test_result$parameter, 1), "\n")
      cat("p-value:", format.pval(test_result$p.value, digits = 3), "\n")
      cat("Mean Difference:", round(test_result$estimate - input$test_value, 3), "\n")
      cat("95% CI for Difference: [", 
          round(test_result$conf.int[1], 3), ", ", 
          round(test_result$conf.int[2], 3), "]\n\n")
    }
  })
  
  # Paired T-Test output
  output$paired_output <- renderPrint({
    req(data(), input$paired_vars)
    
    if (length(input$paired_vars) < 2) return("Please select at least two variables")
    
    conf_level <- input$paired_conf / 100
    alternative <- ifelse(input$paired_tails == "One-tailed", "greater", "two.sided")
    
    cat("Paired Samples T-Test\n")
    cat("=====================\n\n")
    
    # Create pairs of variables
    var_pairs <- matrix(input$paired_vars, ncol = 2, byrow = TRUE)
    
    for (i in 1:nrow(var_pairs)) {
      var1 <- var_pairs[i, 1]
      var2 <- var_pairs[i, 2]
      
      cat("Pair: ", var1, "-", var2, "\n")
      cat("------------------------\n")
      
      # Check if variables are numeric
      if (!is.numeric(data()[[var1]]) || !is.numeric(data()[[var2]])) {
        cat("Skipping non-numeric variables:", var1, "or", var2, "\n\n")
        next
      }
      
      # Perform paired t-test
      test_result <- t.test(
        data()[[var1]], 
        data()[[var2]], 
        paired = TRUE,
        conf.level = conf_level,
        alternative = alternative
      )
      
      # Format output
      cat("t-value:", round(test_result$statistic, 3), "\n")
      cat("df:", round(test_result$parameter, 1), "\n")
      cat("p-value:", format.pval(test_result$p.value, digits = 3), "\n")
      cat("Mean Difference:", round(diff(test_result$estimate), 3), "\n")
      cat("95% CI for Difference: [", 
          round(test_result$conf.int[1], 3), ", ", 
          round(test_result$conf.int[2], 3), "]\n\n")
    }
  })
  
  # Paired T-Test plot
  output$paired_plot <- renderPlotly({
    req(data(), input$paired_vars)
    
    if (length(input$paired_vars) < 2) return()
    
    # Use the first pair for plotting
    var1 <- input$paired_vars[1]
    var2 <- input$paired_vars[2]
    
    # Create a long format data frame for plotting
    plot_data <- data.frame(
      Value = c(data()[[var1]], data()[[var2]]),
      Variable = rep(c(var1, var2), each = nrow(data())),
      ID = rep(1:nrow(data()), 2)
    )
    
    p <- ggplot(plot_data, aes(x = Variable, y = Value, fill = Variable)) +
      geom_boxplot(alpha = 0.7) +
      scale_fill_manual(values = c(spss_blue, "#E27D72")) +
      theme_minimal() +
      labs(title = paste("Boxplot of", var1, "and", var2))
    
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
  
  # Export ANOVA results to Word
  output$export_anova_word <- downloadHandler(
    filename = function() {
      paste("anova_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(data(), input$anova_dv, input$anova_factor)
      
      formula <- as.formula(paste(input$anova_dv, "~", input$anova_factor))
      anova_result <- aov(formula, data = data())
      
      # Create APA-style table
      anova_table <- as.data.frame(summary(anova_result)[[1]])
      anova_table <- cbind(Source = rownames(anova_table), anova_table)
      rownames(anova_table) <- NULL
      colnames(anova_table) <- c("Source", "SS", "df", "MS", "F", "p")
      
      # Create flextable
      ft <- flextable(anova_table)
      ft <- set_caption(ft, caption = "Analysis of Variance")
      ft <- theme_apa(ft)
      ft <- autofit(ft)
      
      # Save to Word
      doc <- read_docx()
      doc <- body_add_flextable(doc, value = ft)
      print(doc, target = file)
    }
  )
  
  # Correlation output
  output$cor_output <- renderPrint({
    req(data(), input$cor_vars)
    
    if (length(input$cor_vars) < 2) return("Please select at least two variables")
    
    cor_data <- data()[, input$cor_vars, drop = FALSE]
    cor_matrix <- cor(cor_data, use = "complete.obs", method = input$cor_method)
    
    # Calculate p-values
    p_matrix <- matrix(NA, nrow = nrow(cor_matrix), ncol = ncol(cor_matrix))
    for (i in 1:nrow(cor_matrix)) {
      for (j in 1:ncol(cor_matrix)) {
        if (i != j) {
          test <- cor.test(cor_data[, i], cor_data[, j], method = input$cor_method,
                           alternative = ifelse(input$sig_test == "One-tailed", "greater", "two.sided"))
          p_matrix[i, j] <- test$p.value
        }
      }
    }
    
    cat("Correlations\n")
    cat("============\n\n")
    
    # Print correlation matrix with significance stars
    cor_df <- as.data.frame(cor_matrix)
    for (i in 1:nrow(cor_df)) {
      for (j in 1:ncol(cor_df)) {
        if (i != j) {
          cor_df[i, j] <- paste0(round(cor_matrix[i, j], 3), 
                                 ifelse(p_matrix[i, j] < 0.001, "***",
                                        ifelse(p_matrix[i, j] < 0.01, "**",
                                               ifelse(p_matrix[i, j] < 0.05, "*", ""))))
        } else {
          cor_df[i, j] <- "-"
        }
      }
    }
    
    print(cor_df)
    
    # Add significance legend
    if (input$flag_sig) {
      cat("\n* p < 0.05, ** p < 0.01, *** p < 0.001\n")
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
  
  # Export correlation results to Word
  output$export_correlation_word <- downloadHandler(
    filename = function() {
      paste("correlation_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(data(), input$cor_vars)
      
      if (length(input$cor_vars) < 2) return()
      
      cor_data <- data()[, input$cor_vars, drop = FALSE]
      cor_matrix <- cor(cor_data, use = "complete.obs", method = input$cor_method)
      
      # Calculate p-values
      p_matrix <- matrix(NA, nrow = nrow(cor_matrix), ncol = ncol(cor_matrix))
      for (i in 1:nrow(cor_matrix)) {
        for (j in 1:ncol(cor_matrix)) {
          if (i != j) {
            test <- cor.test(cor_data[, i], cor_data[, j], method = input$cor_method,
                             alternative = ifelse(input$sig_test == "One-tailed", "greater", "two.sided"))
            p_matrix[i, j] <- test$p.value
          }
        }
      }
      
      # Create a combined matrix with correlations and p-values
      result_matrix <- matrix("", nrow = nrow(cor_matrix), ncol = ncol(cor_matrix))
      for (i in 1:nrow(result_matrix)) {
        for (j in 1:ncol(result_matrix)) {
          if (i == j) {
            result_matrix[i, j] <- "1"
          } else if (i < j) {
            result_matrix[i, j] <- paste0(round(cor_matrix[i, j], 3), 
                                          " (p = ", round(p_matrix[i, j], 3), ")")
          }
        }
      }
      
      # Convert to data frame
      result_df <- as.data.frame(result_matrix)
      colnames(result_df) <- colnames(cor_matrix)
      result_df <- cbind(Variable = rownames(cor_matrix), result_df)
      
      # Create flextable
      ft <- flextable(result_df)
      ft <- set_caption(ft, caption = "Correlation Matrix")
      ft <- theme_apa(ft)
      ft <- autofit(ft)
      
      # Save to Word
      doc <- read_docx()
      doc <- body_add_flextable(doc, value = ft)
      print(doc, target = file)
    }
  )
  
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
  
  # Export regression results to Word
  output$export_regression_word <- downloadHandler(
    filename = function() {
      paste("regression_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(data(), input$reg_dep, input$reg_indep)
      
      if (length(input$reg_indep) == 0) return()
      
      formula <- as.formula(paste(input$reg_dep, "~", paste(input$reg_indep, collapse = "+")))
      model <- lm(formula, data = data())
      
      # Create APA-style table for coefficients
      coef_table <- as.data.frame(summary(model)$coefficients)
      coef_table <- cbind(Variable = rownames(coef_table), coef_table)
      rownames(coef_table) <- NULL
      colnames(coef_table) <- c("Variable", "B", "SE", "t", "p")
      
      # Create flextable
      ft <- flextable(coef_table)
      ft <- set_caption(ft, caption = "Regression Coefficients")
      ft <- theme_apa(ft)
      ft <- autofit(ft)
      
      # Save to Word
      doc <- read_docx()
      doc <- body_add_flextable(doc, value = ft)
      print(doc, target = file)
    }
  )
  
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
  
  # Regression plots
  output$reg_plot <- renderPlot({
    req(data(), input$reg_dep, input$reg_indep)
    
    if (length(input$reg_indep) == 0) return()
    
    formula <- as.formula(paste(input$reg_dep, "~", paste(input$reg_indep, collapse = "+")))
    model <- lm(formula, data = data())
    
    # Create diagnostic plots
    par(mfrow = c(2, 2))
    plot(model)
  })
  
  # Logistic Regression output
  output$logistic_summary <- renderPrint({
    req(data(), input$logistic_dep, input$logistic_indep)
    
    if (length(input$logistic_indep) == 0) return("Please select at least one independent variable")
    
    df <- data()
    dep_var <- df[[input$logistic_dep]]
    
    # Set reference category
    if (!is.null(input$logistic_ref)) {
      if (is.factor(dep_var)) {
        dep_var <- relevel(dep_var, ref = input$logistic_ref)
      } else {
        dep_var <- factor(dep_var)
        dep_var <- relevel(dep_var, ref = as.character(input$logistic_ref))
      }
      df[[input$logistic_dep]] <- dep_var
    }
    
    formula <- as.formula(paste(input$logistic_dep, "~", paste(input$logistic_indep, collapse = "+")))
    model <- glm(formula, data = df, family = binomial)
    
    cat("Logistic Regression\n")
    cat("===================\n\n")
    print(summary(model))
    
    cat("\nOdds Ratios\n")
    cat("===========\n\n")
    odds_ratios <- exp(coef(model))
    print(odds_ratios)
  })
  
  # Export logistic regression results to Word
  output$export_logistic_word <- downloadHandler(
    filename = function() {
      paste("logistic_regression_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(data(), input$logistic_dep, input$logistic_indep)
      
      if (length(input$logistic_indep) == 0) return()
      
      df <- data()
      dep_var <- df[[input$logistic_dep]]
      
      # Set reference category
      if (!is.null(input$logistic_ref)) {
        if (is.factor(dep_var)) {
          dep_var <- relevel(dep_var, ref = input$logistic_ref)
        } else {
          dep_var <- factor(dep_var)
          dep_var <- relevel(dep_var, ref = as.character(input$logistic_ref))
        }
        df[[input$logistic_dep]] <- dep_var
      }
      
      formula <- as.formula(paste(input$logistic_dep, "~", paste(input$logistic_indep, collapse = "+")))
      model <- glm(formula, data = df, family = binomial)
      
      # Create APA-style table for coefficients
      coef_table <- as.data.frame(summary(model)$coefficients)
      coef_table <- cbind(Variable = rownames(coef_table), coef_table)
      rownames(coef_table) <- NULL
      colnames(coef_table) <- c("Variable", "B", "SE", "z", "p")
      
      # Add odds ratios
      coef_table$OR <- exp(coef_table$B)
      coef_table$OR_CI_Lower <- exp(coef_table$B - 1.96 * coef_table$SE)
      coef_table$OR_CI_Upper <- exp(coef_table$B + 1.96 * coef_table$SE)
      
      # Create flextable
      ft <- flextable(coef_table)
      ft <- set_caption(ft, caption = "Logistic Regression Coefficients")
      ft <- theme_apa(ft)
      ft <- autofit(ft)
      
      # Save to Word
      doc <- read_docx()
      doc <- body_add_flextable(doc, value = ft)
      print(doc, target = file)
    }
  )
  
  # Ordinal Regression output
  output$ordinal_output <- renderPrint({
    req(data(), input$ordinal_dep, input$ordinal_indep)
    
    if (length(input$ordinal_indep) == 0) return("Please select at least one independent variable")
    
    # Check if MASS package is available
    if (!requireNamespace("MASS", quietly = TRUE)) {
      return("Please install the MASS package to use ordinal regression")
    }
    
    formula <- as.formula(paste(input$ordinal_dep, "~", paste(input$ordinal_indep, collapse = "+")))
    model <- MASS::polr(formula, data = data(), Hess = TRUE)
    
    cat("Ordinal Regression\n")
    cat("==================\n\n")
    print(summary(model))
  })
  
  # Multinomial Regression output
  output$multinomial_output <- renderPrint({
    req(data(), input$multinomial_dep, input$multinomial_indep)
    
    if (length(input$multinomial_indep) == 0) return("Please select at least one independent variable")
    
    # Check if nnet package is available
    if (!requireNamespace("nnet", quietly = TRUE)) {
      return("Please install the nnet package to use multinomial regression")
    }
    
    df <- data()
    dep_var <- df[[input$multinomial_dep]]
    
    # Set reference category
    if (!is.null(input$multinomial_ref)) {
      if (is.factor(dep_var)) {
        dep_var <- relevel(dep_var, ref = input$multinomial_ref)
      } else {
        dep_var <- factor(dep_var)
        dep_var <- relevel(dep_var, ref = as.character(input$multinomial_ref))
      }
      df[[input$multinomial_dep]] <- dep_var
    }
    
    formula <- as.formula(paste(input$multinomial_dep, "~", paste(input$multinomial_indep, collapse = "+")))
    model <- nnet::multinom(formula, data = df)
    
    cat("Multinomial Regression\n")
    cat("======================\n\n")
    print(summary(model))
  })
  
  # Poisson Regression output
  output$poisson_output <- renderPrint({
    req(data(), input$poisson_dep, input$poisson_indep)
    
    if (length(input$poisson_indep) == 0) return("Please select at least one independent variable")
    
    formula <- as.formula(paste(input$poisson_dep, "~", paste(input$poisson_indep, collapse = "+")))
    model <- glm(formula, data = data(), family = poisson)
    
    cat("Poisson Regression\n")
    cat("==================\n\n")
    print(summary(model))
  })
  
  # Factor Analysis output
  output$factor_summary <- renderPrint({
    req(data(), input$factor_vars, input$num_factors)
    
    if (length(input$factor_vars) < 2) return("Please select at least two variables")
    
    fa_data <- data()[, input$factor_vars, drop = FALSE]
    fa_data <- na.omit(fa_data)
    
    # Perform factor analysis
    fa_result <- factanal(fa_data, factors = input$num_factors, rotation = input$rotation_method)
    
    cat("Factor Analysis\n")
    cat("===============\n\n")
    print(fa_result, digits = 3, sort = TRUE)
  })
  
  # Factor loadings
  output$factor_loadings <- renderPrint({
    req(data(), input$factor_vars, input$num_factors)
    
    if (length(input$factor_vars) < 2) return()
    
    fa_data <- data()[, input$factor_vars, drop = FALSE]
    fa_data <- na.omit(fa_data)
    
    # Perform factor analysis
    fa_result <- factanal(fa_data, factors = input$num_factors, rotation = input$rotation_method)
    
    cat("Factor Loadings\n")
    cat("===============\n\n")
    print(fa_result$loadings, digits = 3, cutoff = 0.3)
  })
  
  # Factor scree plot
  output$factor_scree <- renderPlotly({
    req(data(), input$factor_vars)
    
    if (length(input$factor_vars) < 2) return()
    
    fa_data <- data()[, input$factor_vars, drop = FALSE]
    fa_data <- na.omit(fa_data)
    
    # Calculate eigenvalues
    cor_matrix <- cor(fa_data)
    eigenvalues <- eigen(cor_matrix)$values
    
    # Create scree plot
    scree_data <- data.frame(
      Factor = 1:length(eigenvalues),
      Eigenvalue = eigenvalues
    )
    
    p <- ggplot(scree_data, aes(x = Factor, y = Eigenvalue)) +
      geom_line(color = spss_blue) +
      geom_point(color = spss_blue, size = 2) +
      geom_hline(yintercept = 1, linetype = "dashed", color = "red") +
      theme_minimal() +
      labs(title = "Scree Plot", x = "Factor Number", y = "Eigenvalue")
    
    ggplotly(p)
  })
  
  # CFA output
  output$cfa_summary <- renderPrint({
    req(data(), input$cfa_vars, input$cfa_model)
    
    if (length(input$cfa_vars) < 2) return("Please select at least two variables")
    if (input$cfa_model == "") return("Please specify a model")
    
    cfa_data <- data()[, input$cfa_vars, drop = FALSE]
    cfa_data <- na.omit(cfa_data)
    
    # Perform CFA
    cfa_result <- cfa(input$cfa_model, data = cfa_data)
    
    cat("Confirmatory Factor Analysis\n")
    cat("===========================\n\n")
    summary(cfa_result, fit.measures = TRUE, standardized = TRUE)
  })
  
  # Reliability Analysis output
  output$reliability_output <- renderPrint({
    req(data(), input$reliability_vars)
    
    if (length(input$reliability_vars) < 2) return("Please select at least two variables")
    
    rel_data <- data()[, input$reliability_vars, drop = FALSE]
    rel_data <- na.omit(rel_data)
    
    cat("Reliability Analysis\n")
    cat("====================\n\n")
    
    # Calculate Cronbach's alpha
    if ("Cronbach's alpha" %in% input$reliability_stats) {
      alpha_result <- psych::alpha(rel_data)
      cat("Cronbach's Alpha:\n")
      print(alpha_result$total$raw_alpha)
      cat("\n")
    }
    
    # Calculate McDonald's omega
    if ("McDonald's omega" %in% input$reliability_stats) {
      omega_result <- omega(rel_data, plot = FALSE)
      cat("McDonald's Omega:\n")
      print(omega_result$omega.tot)
      cat("\n")
    }
    
    # Item statistics
    if ("Item statistics" %in% input$reliability_stats) {
      cat("Item Statistics:\n")
      item_stats <- describe(rel_data)
      print(item_stats)
      cat("\n")
    }
    
    # Scale statistics
    if ("Scale statistics" %in% input$reliability_stats) {
      cat("Scale Statistics:\n")
      scale_stats <- describe(rowSums(rel_data, na.rm = TRUE))
      print(scale_stats)
      cat("\n")
    }
  })
  
  # Chi-Square output
  output$chi_crosstab <- renderPrint({
    req(data(), input$chi_row, input$chi_col)
    
    df <- data()
    row_var <- df[[input$chi_row]]
    col_var <- df[[input$chi_col]]
    
    # Create cross tabulation
    crosstab <- table(row_var, col_var)
    
    cat("Crosstabulation\n")
    cat("===============\n\n")
    print(crosstab)
    
    cat("\nRow Percentages\n")
    cat("===============\n\n")
    print(prop.table(crosstab, 1) * 100)
    
    cat("\nColumn Percentages\n")
    cat("==================\n\n")
    print(prop.table(crosstab, 2) * 100)
  })
  
  output$chi_test <- renderPrint({
    req(data(), input$chi_row, input$chi_col)
    
    df <- data()
    row_var <- df[[input$chi_row]]
    col_var <- df[[input$chi_col]]
    
    # Perform chi-square test
    chi_test <- chisq.test(row_var, col_var)
    
    cat("Chi-Square Test\n")
    cat("===============\n\n")
    print(chi_test)
  })
  
  # Chi-Square plot
  output$chi_plot <- renderPlotly({
    req(data(), input$chi_row, input$chi_col)
    
    df <- data()
    row_var <- df[[input$chi_row]]
    col_var <- df[[input$chi_col]]
    
    # Create a stacked bar chart
    plot_data <- as.data.frame(table(row_var, col_var))
    colnames(plot_data) <- c("Row", "Column", "Frequency")
    
    p <- ggplot(plot_data, aes(x = Row, y = Frequency, fill = Column)) +
      geom_bar(stat = "identity", position = "stack") +
      scale_fill_brewer(palette = "Set1") +
      theme_minimal() +
      labs(title = paste("Crosstabulation of", input$chi_row, "and", input$chi_col))
    
    ggplotly(p)
  })
  
  # Mann-Whitney U Test output
  output$mw_output <- renderPrint({
    req(data(), input$mw_vars, input$mw_group_var, input$mw_group_def)
    
    # Parse group definition
    groups <- trimws(unlist(strsplit(input$mw_group_def, ",")))
    if (length(groups) != 2) {
      return("Please define exactly two groups separated by a comma (e.g., '1,2' or 'A,B')")
    }
    
    # Filter data based on groups
    group_var <- data()[[input$mw_group_var]]
    if (!all(groups %in% unique(group_var))) {
      return("One or both group values not found in the grouping variable")
    }
    
    # Create subset for each group
    group1_data <- data()[group_var == groups[1], input$mw_vars, drop = FALSE]
    group2_data <- data()[group_var == groups[2], input$mw_vars, drop = FALSE]
    
    cat("Mann-Whitney U Test\n")
    cat("===================\n\n")
    
    for (var in input$mw_vars) {
      cat("Variable: ", var, "\n")
      cat("------------------------\n")
      
      # Check if variable is numeric
      if (!is.numeric(data()[[var]])) {
        cat("Skipping non-numeric variable:", var, "\n\n")
        next
      }
      
      # Perform Mann-Whitney U test
      test_result <- wilcox.test(
        group1_data[[var]], 
        group2_data[[var]]
      )
      
      # Format output
      cat("W-value:", round(test_result$statistic, 3), "\n")
      cat("p-value:", format.pval(test_result$p.value, digits = 3), "\n\n")
    }
  })
  
  # Kruskal-Wallis Test output
  output$kw_output <- renderPrint({
    req(data(), input$kw_vars, input$kw_group_var)
    
    cat("Kruskal-Wallis Test\n")
    cat("===================\n\n")
    
    for (var in input$kw_vars) {
      cat("Variable: ", var, "\n")
      cat("------------------------\n")
      
      # Check if variable is numeric
      if (!is.numeric(data()[[var]])) {
        cat("Skipping non-numeric variable:", var, "\n\n")
        next
      }
      
      # Perform Kruskal-Wallis test
      formula <- as.formula(paste(var, "~", input$kw_group_var))
      test_result <- kruskal.test(formula, data = data())
      
      # Format output
      cat("Chi-squared:", round(test_result$statistic, 3), "\n")
      cat("df:", test_result$parameter, "\n")
      cat("p-value:", format.pval(test_result$p.value, digits = 3), "\n\n")
    }
  })
  
  # Wilcoxon Signed-Rank Test output
  output$wilcoxon_output <- renderPrint({
    req(data(), input$wilcoxon_vars)
    
    if (length(input$wilcoxon_vars) < 2) return("Please select at least two variables")
    
    cat("Wilcoxon Signed-Rank Test\n")
    cat("=========================\n\n")
    
    # Create pairs of variables
    var_pairs <- matrix(input$wilcoxon_vars, ncol = 2, byrow = TRUE)
    
    for (i in 1:nrow(var_pairs)) {
      var1 <- var_pairs[i, 1]
      var2 <- var_pairs[i, 2]
      
      cat("Pair: ", var1, "-", var2, "\n")
      cat("------------------------\n")
      
      # Check if variables are numeric
      if (!is.numeric(data()[[var1]]) || !is.numeric(data()[[var2]])) {
        cat("Skipping non-numeric variables:", var1, "or", var2, "\n\n")
        next
      }
      
      # Perform Wilcoxon signed-rank test
      test_result <- wilcox.test(
        data()[[var1]], 
        data()[[var2]], 
        paired = TRUE
      )
      
      # Format output
      cat("V-value:", round(test_result$statistic, 3), "\n")
      cat("p-value:", format.pval(test_result$p.value, digits = 3), "\n\n")
    }
  })
  
  # Chart output
  output$chart_output <- renderPlot({
    req(data(), input$chart_vars, input$generate_chart)
    
    if (input$chart_type == "Bar") {
      req(length(input$chart_vars) >= 1)
      var <- input$chart_vars[1]
      
      if (input$chart_group != "None") {
        p <- ggplot(data(), aes_string(x = input$chart_group, fill = input$chart_group)) +
          geom_bar(alpha = 0.7) +
          labs(title = ifelse(input$chart_title != "", input$chart_title, paste("Bar Chart of", var)))
      } else {
        p <- ggplot(data(), aes_string(x = var)) +
          geom_bar(fill = spss_blue, alpha = 0.7) +
          labs(title = ifelse(input$chart_title != "", input$chart_title, paste("Bar Chart of", var)))
      }
    } else if (input$chart_type == "Histogram") {
      req(length(input$chart_vars) >= 1)
      var <- input$chart_vars[1]
      
      p <- ggplot(data(), aes_string(x = var)) +
        geom_histogram(fill = spss_blue, alpha = 0.7, bins = 30) +
        labs(title = ifelse(input$chart_title != "", input$chart_title, paste("Histogram of", var)))
    } else if (input$chart_type == "Boxplot") {
      req(length(input$chart_vars) >= 1)
      var <- input$chart_vars[1]
      
      if (input$chart_group != "None") {
        p <- ggplot(data(), aes_string(x = input$chart_group, y = var, fill = input$chart_group)) +
          geom_boxplot(alpha = 0.7) +
          labs(title = ifelse(input$chart_title != "", input$chart_title, paste("Boxplot of", var, "by", input$chart_group)))
      } else {
        p <- ggplot(data(), aes_string(y = var)) +
          geom_boxplot(fill = spss_blue, alpha = 0.7) +
          labs(title = ifelse(input$chart_title != "", input$chart_title, paste("Boxplot of", var)))
      }
    } else if (input$chart_type == "Scatterplot") {
      req(length(input$chart_vars) >= 2)
      x_var <- input$chart_vars[1]
      y_var <- input$chart_vars[2]
      
      if (input$chart_group != "None") {
        p <- ggplot(data(), aes_string(x = x_var, y = y_var, color = input$chart_group)) +
          geom_point(alpha = 0.7) +
          labs(title = ifelse(input$chart_title != "", input$chart_title, paste("Scatterplot of", y_var, "vs", x_var)))
      } else {
        p <- ggplot(data(), aes_string(x = x_var, y = y_var)) +
          geom_point(color = spss_blue, alpha = 0.7) +
          labs(title = ifelse(input$chart_title != "", input$chart_title, paste("Scatterplot of", y_var, "vs", x_var)))
      }
    } else if (input$chart_type == "Density") {
      req(length(input$chart_vars) >= 1)
      var <- input$chart_vars[1]
      
      if (input$chart_group != "None") {
        p <- ggplot(data(), aes_string(x = var, fill = input$chart_group)) +
          geom_density(alpha = 0.5) +
          labs(title = ifelse(input$chart_title != "", input$chart_title, paste("Density Plot of", var)))
      } else {
        p <- ggplot(data(), aes_string(x = var)) +
          geom_density(fill = spss_blue, alpha = 0.7) +
          labs(title = ifelse(input$chart_title != "", input$chart_title, paste("Density Plot of", var)))
      }
    } else if (input$chart_type == "Violin") {
      req(length(input$chart_vars) >= 1)
      var <- input$chart_vars[1]
      
      if (input$chart_group != "None") {
        p <- ggplot(data(), aes_string(x = input$chart_group, y = var, fill = input$chart_group)) +
          geom_violin(alpha = 0.7) +
          labs(title = ifelse(input$chart_title != "", input$chart_title, paste("Violin Plot of", var, "by", input$chart_group)))
      } else {
        p <- ggplot(data(), aes_string(x = 1, y = var)) +
          geom_violin(fill = spss_blue, alpha = 0.7) +
          labs(title = ifelse(input$chart_title != "", input$chart_title, paste("Violin Plot of", var))) +
          theme(axis.title.x = element_blank(), axis.text.x = element_blank())
      }
    } else if (input$chart_type == "Q-Q Plot") {
      req(length(input$chart_vars) >= 1)
      var <- input$chart_vars[1]
      
      p <- ggplot(data(), aes_string(sample = var)) +
        stat_qq() +
        stat_qq_line() +
        labs(title = ifelse(input$chart_title != "", input$chart_title, paste("Q-Q Plot of", var)))
    } else {
      return()
    }
    
    # Apply color palette
    if (input$color_palette == "spss") {
      p <- p + scale_fill_manual(values = c(spss_blue, "#E27D72", "#68C3A3", "#F7C46C"))
    } else {
      p <- p + scale_fill_brewer(palette = input$color_palette)
    }
    
    # Apply theme and text sizes
    p <- p + theme_minimal() +
      theme(
        plot.title = element_text(size = input$title_size),
        axis.text = element_text(size = input$axis_size),
        axis.title = element_text(size = input$axis_size)
      )
    
    p
  })
  
  # Download chart as PNG
  output$download_chart <- downloadHandler(
    filename = function() {
      paste("chart_", Sys.Date(), ".png", sep = "")
    },
    content = function(file) {
      req(data(), input$chart_vars, input$generate_chart)
      
      if (input$chart_type == "Bar") {
        req(length(input$chart_vars) >= 1)
        var <- input$chart_vars[1]
        
        if (input$chart_group != "None") {
          p <- ggplot(data(), aes_string(x = input$chart_group, fill = input$chart_group)) +
            geom_bar(alpha = 0.7) +
            labs(title = ifelse(input$chart_title != "", input$chart_title, paste("Bar Chart of", var)))
        } else {
          p <- ggplot(data(), aes_string(x = var)) +
            geom_bar(fill = spss_blue, alpha = 0.7) +
            labs(title = ifelse(input$chart_title != "", input$chart_title, paste("Bar Chart of", var)))
        }
      } else if (input$chart_type == "Histogram") {
        req(length(input$chart_vars) >= 1)
        var <- input$chart_vars[1]
        
        p <- ggplot(data(), aes_string(x = var)) +
          geom_histogram(fill = spss_blue, alpha = 0.7, bins = 30) +
          labs(title = ifelse(input$chart_title != "", input$chart_title, paste("Histogram of", var)))
      } else if (input$chart_type == "Boxplot") {
        req(length(input$chart_vars) >= 1)
        var <- input$chart_vars[1]
        
        if (input$chart_group != "None") {
          p <- ggplot(data(), aes_string(x = input$chart_group, y = var, fill = input$chart_group)) +
            geom_boxplot(alpha = 0.7) +
            labs(title = ifelse(input$chart_title != "", input$chart_title, paste("Boxplot of", var, "by", input$chart_group)))
        } else {
          p <- ggplot(data(), aes_string(y = var)) +
            geom_boxplot(fill = spss_blue, alpha = 0.7) +
            labs(title = ifelse(input$chart_title != "", input$chart_title, paste("Boxplot of", var)))
        }
      } else if (input$chart_type == "Scatterplot") {
        req(length(input$chart_vars) >= 2)
        x_var <- input$chart_vars[1]
        y_var <- input$chart_vars[2]
        
        if (input$chart_group != "None") {
          p <- ggplot(data(), aes_string(x = x_var, y = y_var, color = input$chart_group)) +
            geom_point(alpha = 0.7) +
            labs(title = ifelse(input$chart_title != "", input$chart_title, paste("Scatterplot of", y_var, "vs", x_var)))
        } else {
          p <- ggplot(data(), aes_string(x = x_var, y = y_var)) +
            geom_point(color = spss_blue, alpha = 0.7) +
            labs(title = ifelse(input$chart_title != "", input$chart_title, paste("Scatterplot of", y_var, "vs", x_var)))
        }
      } else if (input$chart_type == "Density") {
        req(length(input$chart_vars) >= 1)
        var <- input$chart_vars[1]
        
        if (input$chart_group != "None") {
          p <- ggplot(data(), aes_string(x = var, fill = input$chart_group)) +
            geom_density(alpha = 0.5) +
            labs(title = ifelse(input$chart_title != "", input$chart_title, paste("Density Plot of", var)))
        } else {
          p <- ggplot(data(), aes_string(x = var)) +
            geom_density(fill = spss_blue, alpha = 0.7) +
            labs(title = ifelse(input$chart_title != "", input$chart_title, paste("Density Plot of", var)))
        }
      } else if (input$chart_type == "Violin") {
        req(length(input$chart_vars) >= 1)
        var <- input$chart_vars[1]
        
        if (input$chart_group != "None") {
          p <- ggplot(data(), aes_string(x = input$chart_group, y = var, fill = input$chart_group)) +
            geom_violin(alpha = 0.7) +
            labs(title = ifelse(input$chart_title != "", input$chart_title, paste("Violin Plot of", var, "by", input$chart_group)))
        } else {
          p <- ggplot(data(), aes_string(x = 1, y = var)) +
            geom_violin(fill = spss_blue, alpha = 0.7) +
            labs(title = ifelse(input$chart_title != "", input$chart_title, paste("Violin Plot of", var))) +
            theme(axis.title.x = element_blank(), axis.text.x = element_blank())
        }
      } else if (input$chart_type == "Q-Q Plot") {
        req(length(input$chart_vars) >= 1)
        var <- input$chart_vars[1]
        
        p <- ggplot(data(), aes_string(sample = var)) +
          stat_qq() +
          stat_qq_line() +
          labs(title = ifelse(input$chart_title != "", input$chart_title, paste("Q-Q Plot of", var)))
      } else {
        return()
      }
      
      # Apply color palette
      if (input$color_palette == "spss") {
        p <- p + scale_fill_manual(values = c(spss_blue, "#E27D72", "#68C3A3", "#F7C46C"))
      } else {
        p <- p + scale_fill_brewer(palette = input$color_palette)
      }
      
      # Apply theme and text sizes
      p <- p + theme_minimal() +
        theme(
          plot.title = element_text(size = input$title_size),
          axis.text = element_text(size = input$axis_size),
          axis.title = element_text(size = input$axis_size)
        )
      
      # Save the plot
      ggsave(file, plot = p, width = input$chart_width/100, height = input$chart_height/100, dpi = 300)
    }
  )
  
  # Compute variable functionality
  observeEvent(input$compute_var, {
    req(data(), input$target_var, input$numeric_expr)
    
    df <- data()
    tryCatch({
      # Parse and evaluate the expression
      expr <- parse(text = input$numeric_expr)
      df[[input$target_var]] <- eval(expr, envir = df)
      data(df)
      data_modified(TRUE)
      
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
      update_all_select_inputs(df, types, session)
      
      showNotification(paste("Variable", input$target_var, "created successfully"), type = "message")
    }, error = function(e) {
      showNotification(paste("Error:", e$message), type = "error")
    })
  })
}

# Run the application
shinyApp(ui = ui, server = server)
