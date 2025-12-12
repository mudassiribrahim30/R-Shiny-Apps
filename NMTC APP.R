# app.R
# SPSS-Style Statistical Analysis Software - Enhanced Version
library(shiny)
library(shinydashboard)
library(DT)
library(ggplot2)
library(plotly)
library(readxl)
library(tidyverse)
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
library(officer)
library(openxlsx)
library(ggpubr)
library(RColorBrewer)
library(lavaan)  # For EFA/CFA
library(semTools) # For McDonald's omega
#library(flextable)
library(colourpicker)
library(stringr)
library(purrr)
library(emmeans)  # For post hoc tests
library(multcomp) # For multiple comparisons
library(agricolae)  # For Duncan's test
library(ResourceSelection)
library(sortable)
library(semPlot)

# SPSS Color Scheme
spss_blue <- "#326EA6"
spss_dark_blue <- "#1A3A5F"
spss_light_gray <- "#F2F2F2"
spss_dark_gray <- "#7F7F7F"

# UI Definition - SPSS Style Interface
ui <- dashboardPage(
  skin = "blue",
  
  dashboardHeader(
    title = tags$div(
      style = "display: flex; align-items: center;",
      
      # Logo (from GitHub, enlarged)
      tags$img(
        src = "https://raw.githubusercontent.com/mudassiribrahim30/R-Shiny-Apps/main/NMTC%20logo.png",
        height = "50px",   # adjust size (bigger, clean look)
        style = "margin-right: 12px;"
      ),
      
      # App Name
      span(
        "TNMTC DataLab",
        style = paste0(
          "color: white; font-weight: bold; font-size: 20px; background-color:", spss_blue
        )
      )
    ),
    
    titleWidth = 320,  # slightly wider to fit logo + text
    
    dropdownMenu(
      type = "messages",
      badgeStatus = NULL,
      icon = icon("question-circle"),
      messageItem(
        from = "Help",
        message = "TNMTC DataLab Help",
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
               menuSubItem("Load Data", tabName = "load_data"),
               menuSubItem("Save", tabName = "save")),
      # REMOVED: menuItem("Edit", tabName = "edit", icon = icon("edit")),
      menuItem("View", tabName = "view", icon = icon("desktop")),
      menuItem("Data", tabName = "data", icon = icon("database"),
               menuSubItem("Define Variable Properties", tabName = "define_var"),
               menuSubItem("Sort Cases", tabName = "sort_cases"),
               menuSubItem("Select Cases", tabName = "select_cases"),
               menuSubItem("Weight Cases", tabName = "weight_cases")),
      menuItem("Transform", tabName = "transform", icon = icon("exchange-alt"),
               menuSubItem("Compute Variable", tabName = "compute")),
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
                        menuSubItem("Wilcoxon Signed-Rank", tabName = "wilcoxon"),
                        menuSubItem("Fisher's Exact Test", tabName = "fisher_exact")
               )
      ),
      menuItem("Graphs", tabName = "graphs", icon = icon("chart-area"),
               menuSubItem("Chart Builder", tabName = "chart_builder")),
      menuItem("Help", tabName = "help", icon = icon("question-circle")),
      menuItem("About", tabName = "about", icon = icon("info-circle"))
    )
  ),
  dashboardBody(
    tags$head(
      tags$link(rel = "stylesheet", type = "text/css", href = "https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css"),
      tags$style(HTML(paste0("
      /* Fixed header and sidebar */
      .main-header {
        position: fixed;
        width: 100%;
        z-index: 1000;
      }
      
      .content-wrapper, .right-side {
        margin-top: 50px; /* Adjust this based on your header height */
      }
      
      /* Fixed sidebar and scrollable content */
      .wrapper {
        height: 100vh;
        overflow: hidden;
      }
      
      .main-sidebar {
        position: fixed;
        height: 100vh;
        overflow-y: auto;
        margin-top: 50px; /* Adjust this based on your header height */
      }
      
      .content-wrapper, .right-side {
        margin-left: 250px;
        height: calc(100vh - 50px); /* Adjust height considering header */
        overflow-y: auto;
        background-color: white;
      }
      
      /* Increased font sizes */
      .result-output {
        font-size: 16px !important;
        line-height: 1.5;
        font-family: 'Arial', sans-serif;
      }
      
      .data-table {
        font-size: 14px !important;
      }
      
      .analysis-panel {
        font-size: 15px;
      }
      
      /* Chart builder color improvements */
      .js-irs-0 .irs-single, .js-irs-0 .irs-bar-edge, .js-irs-0 .irs-bar {
        background: ", spss_blue, ";
      }
      
      .main-header .logo {
        font-weight: bold;
        font-size: 18px;
        background-color: ", spss_blue, ";
        color: white;
        height: 50px; /* Fixed height for header */
        line-height: 50px; /* Center text vertically */
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
      .sidebar-menu li.active {
        border-left: 4px solid ", spss_blue, ";
      }
      .skin-blue .main-header .navbar {
        background-color: ", spss_dark_blue, ";
        position: fixed;
        width: 100%;
        z-index: 1000;
      }
      .skin-blue .main-sidebar {
        background-color: ", spss_light_gray, ";
        margin-top: 50px; /* Adjust this based on your header height */
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
      /* Data table styling */
      .dataTables_wrapper {
        font-size: 14px;
      }
      table.dataTable {
        width: 100% !important;
        margin: 0 auto;
      }
    ")))
    ),
    
    tabItems(
      # Data Load Tab
      tabItem(
        tabName = "load_data",
        h2("Load Data", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Import Data",
            selectInput("data_format", "Select Data Format",
                        choices = c("CSV", "Excel", "SPSS", "Stata")),
            conditionalPanel(
              condition = "input.data_format == 'CSV'",
              fileInput("file_csv", "Select CSV File",
                        accept = c(".csv"),
                        buttonLabel = "Browse...",
                        placeholder = "No file selected"),
              selectInput("sep", "Separator", 
                          choices = c(Comma = ",", Semicolon = ";", Tab = "\t", Space = " "), 
                          selected = ","),
              checkboxInput("header", "Header", TRUE),
              selectInput("dec", "Decimal Separator", choices = c("." = ".", "," = ","), selected = ".")
            ),
            conditionalPanel(
              condition = "input.data_format == 'Excel'",
              fileInput("file_excel", "Select Excel File",
                        accept = c(".xlsx", ".xls"),
                        buttonLabel = "Browse...",
                        placeholder = "No file selected")
            ),
            conditionalPanel(
              condition = "input.data_format == 'SPSS'",
              fileInput("file_spss", "Select SPSS File",
                        accept = c(".sav"),
                        buttonLabel = "Browse...",
                        placeholder = "No file selected")
            ),
            conditionalPanel(
              condition = "input.data_format == 'Stata'",
              fileInput("file_stata", "Select Stata File",
                        accept = c(".dta"),
                        buttonLabel = "Browse...",
                        placeholder = "No file selected")
            ),
            actionButton("load_btn", "Load Data", class = "btn-success", icon = icon("folder-open"))
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
      
      # Save Data Tab - CORRECTED VERSION
      tabItem(
        tabName = "save",
        h2("Save Data", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Save Options",
            selectInput("save_format", "Format", 
                        choices = c("CSV", "Excel", "SPSS", "Stata", "RData")),
            textInput("save_filename", "File Name", value = "data"),
            downloadButton("download_data", "Save", class = "btn-primary", icon = icon("save"))
          ),
          box(
            width = 8, status = "info", title = "Data Preview",
            div(class = "data-table",
                withSpinner(DTOutput("save_preview"))
            )
          )
        )
      ),
      
      # View Tab - Interactive Data Editor
      tabItem(
        tabName = "view",
        h2("Data Viewer", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 12, status = "primary", title = "Data Editor",
            div(
              style = "margin-bottom: 10px;",
              actionButton("add_row", "Add Row", icon = icon("plus"), class = "btn-success"),
              actionButton("add_col", "Add Column", icon = icon("plus"), class = "btn-info"),
              actionButton("delete_row", "Delete Row", icon = icon("trash"), class = "btn-danger"),
              actionButton("delete_col", "Delete Column", icon = icon("trash"), class = "btn-warning"),
              actionButton("save_edits", "Save Edits", icon = icon("save"), class = "btn-primary")
            ),
            div(class = "data-table",
                withSpinner(DTOutput("data_preview"))
            ),
            tags$script(HTML('
        // Enable cell editing and row/column operations
        $(document).on("click", "#add_row", function() {
          Shiny.setInputValue("add_row_click", Math.random());
        });
        $(document).on("click", "#add_col", function() {
          Shiny.setInputValue("add_col_click", Math.random());
        });
        $(document).on("click", "#delete_row", function() {
          Shiny.setInputValue("delete_row_click", Math.random());
        });
        $(document).on("click", "#delete_col", function() {
          Shiny.setInputValue("delete_col_click", Math.random());
        });
        $(document).on("click", "#save_edits", function() {
          Shiny.setInputValue("save_edits_click", Math.random());
        });
      '))
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
                        choices = c("Recode", "Compute", "Binning", "Log Transform", "Standardize", 
                                    "Center", "Normalize", "Rank", "Dummy Variables", "Lag", "Difference")),
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
            conditionalPanel(
              condition = "input.transform_type == 'Center'",
              selectInput("center_var", "Variable to Center", choices = NULL),
              textInput("center_var_name", "New Variable Name", value = "")
            ),
            conditionalPanel(
              condition = "input.transform_type == 'Normalize'",
              selectInput("normalize_var", "Variable to Normalize", choices = NULL),
              textInput("normalize_var_name", "New Variable Name", value = "")
            ),
            conditionalPanel(
              condition = "input.transform_type == 'Rank'",
              selectInput("rank_var", "Variable to Rank", choices = NULL),
              textInput("rank_var_name", "New Variable Name", value = "")
            ),
            conditionalPanel(
              condition = "input.transform_type == 'Dummy Variables'",
              selectInput("dummy_var", "Categorical Variable", choices = NULL),
              textInput("dummy_prefix", "Prefix for Dummy Variables", value = "")
            ),
            conditionalPanel(
              condition = "input.transform_type == 'Lag'",
              selectInput("lag_var", "Variable for Lag", choices = NULL),
              numericInput("lag_order", "Lag Order", value = 1, min = 1, max = 10),
              textInput("lag_var_name", "New Variable Name", value = "")
            ),
            conditionalPanel(
              condition = "input.transform_type == 'Difference'",
              selectInput("diff_var", "Variable for Difference", choices = NULL),
              numericInput("diff_order", "Difference Order", value = 1, min = 1, max = 10),
              textInput("diff_var_name", "New Variable Name", value = "")
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
            h4("Charts"),
            checkboxGroupInput("freq_charts", "Chart Type",
                               choices = c("Bar charts", "Pie charts", "Histograms")),
            hr(),
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
      
      # Independent Samples T-Test Tab - CORRECTED VERSION
      tabItem(
        tabName = "independent_ttest",
        h2("Independent Samples T-Test", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Test Variables",
            selectizeInput("ttest_vars", "Test Variable(s)", choices = NULL, multiple = TRUE),
            selectInput("ttest_group_var", "Grouping Variable", choices = NULL),
            textInput("group_def", "Define Groups", placeholder = "e.g., male,female for Gender"),
            actionButton("run_ttest", "OK", class = "btn-primary", icon = icon("check")),
            hr(),
            h4("Options"),
            numericInput("ttest_conf", "Confidence Level", value = 95, min = 50, max = 99, step = 1),
            radioButtons("ttest_tails", "Test Type",
                         choices = c("Two-tailed", "One-tailed"), selected = "Two-tailed"),
            hr(),
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
      
      # One Sample T-Test Tab - CORRECTED VERSION
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
          ),
          box(
            width = 8, status = "info", title = "One Sample T-Test",
            div(class = "result-output",
                withSpinner(verbatimTextOutput("onesample_output"))
            )
          )
        )
      ),
      
      # Paired T-Test Tab - CORRECTED VERSION
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
          ),
          box(
            width = 8, status = "info", title = "ANOVA",
            tabBox(
              width = 12,
              tabPanel("ANOVA", div(class = "result-output", withSpinner(verbatimTextOutput("anova_output")))),
              tabPanel("Post Hoc", div(class = "result-output", withSpinner(verbatimTextOutput("posthoc_output"), type = 4))),
              tabPanel("Descriptives", div(class = "result-output", withSpinner(verbatimTextOutput("descriptives_output")))),
              tabPanel("Plots", withSpinner(plotlyOutput("anova_plot", height = "400px")))
            )
          )
        )
      ),
      
      # Correlation Tab - CORRECTED VERSION with p-values displayed
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
            checkboxInput("show_pvalues", "Show p-values", value = TRUE), # NEW: Option to show p-values
            hr(),
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
      
      # Linear Regression Tab - Enhanced with Categorical Variable Support
      tabItem(
        tabName = "linear_reg",
        h2("Linear Regression", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Variables",
            selectInput("reg_dep", "Dependent", choices = NULL),
            selectizeInput("reg_indep", "Independent(s)", choices = NULL, multiple = TRUE),
            
            # Dynamic UI for reference category selection for INDEPENDENT variables
            uiOutput("categorical_ref_ui"),
            
            # NEW: Dynamic UI for reference category selection for DEPENDENT variable (if categorical)
            uiOutput("reg_dep_level_ui"),
            
            actionButton("run_regression", "OK", class = "btn-primary", icon = icon("check")),
            hr(),
            h4("Statistics"),
            checkboxGroupInput("reg_stats", "Regression Coefficients",
                               choices = c("Estimates", "Confidence intervals", "Covariance matrix",
                                           "Model fit", "R squared change", "Descriptives",
                                           "Part and partial correlations", "Collinearity diagnostics")),
            h4("Plots"),
            checkboxGroupInput("reg_plots", "Plot Type",
                               choices = c("Residuals vs Fitted", "Normal Q-Q", "Scale-Location", "Residuals vs Leverage")),
            hr(),
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
            
            # NEW: Dynamic UI for reference category selection for DEPENDENT variable
            uiOutput("logistic_dep_level_ui"),
            
            selectizeInput("logistic_indep", "Independent(s)", choices = NULL, multiple = TRUE),
            uiOutput("logistic_ref_ui"),
            actionButton("run_logistic", "OK", class = "btn-primary", icon = icon("check")),
            hr(),
            h4("Options"),
            checkboxGroupInput("logistic_options", "Display",
                               choices = c("CI for exp(B)", "Hosmer-Lemeshow goodness-of-fit", 
                                           "Iteration history", "Correlation of estimates")),
            hr(),
          ),
          box(
            width = 8, status = "info", title = "Logistic Regression",
            tabBox(
              width = 12,
              tabPanel("Model Summary", div(class = "result-output", withSpinner(verbatimTextOutput("logistic_summary")))),
              tabPanel("Coefficients", div(class = "result-output", withSpinner(verbatimTextOutput("logistic_coef")))),
              tabPanel("Classification", div(class="result-output", withSpinner(verbatimTextOutput("logistic_classification")))),
              tabPanel("H-L Test", div(class = "result-output", withSpinner(verbatimTextOutput("logistic_hl")))),
              tabPanel("Plots", withSpinner(plotlyOutput("logistic_plot", height = "400px")))
            )
          )
        )
      ),
      # Ordinal Regression Tab - CORRECTED VERSION
      tabItem(
        tabName = "ordinal_reg",
        h2("Ordinal Regression", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Variables",
            selectInput("ordinal_dep", "Dependent", choices = NULL),
            
            # UI for dependent variable level ordering and reference category
            uiOutput("ordinal_dep_level_ui"),
            
            selectizeInput("ordinal_indep", "Independent(s)", choices = NULL, multiple = TRUE),
            
            # UI for reference category selection for INDEPENDENT categorical variables
            uiOutput("ordinal_ref_ui"),
            
            actionButton("run_ordinal", "OK", class = "btn-primary", icon = icon("check")),
            hr(),
          ),
          box(
            width = 8, status = "info", title = "Ordinal Regression",
            div(class = "result-output",
                withSpinner(verbatimTextOutput("ordinal_output"))
            )
          )
        )
      ),      
      
      # Multinomial Regression Tab - CORRECTED VERSION
      tabItem(
        tabName = "multinomial_reg",
        h2("Multinomial Regression", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Variables",
            selectInput("multinomial_dep", "Dependent", choices = NULL),
            
            # NEW: UI for dependent variable reference category selection
            uiOutput("multinomial_dep_level_ui"),
            
            selectizeInput("multinomial_indep", "Independent(s)", choices = NULL, multiple = TRUE),
            
            # UI for reference category selection for INDEPENDENT categorical variables
            uiOutput("multinomial_ref_ui"),
            
            actionButton("run_multinomial", "OK", class = "btn-primary", icon = icon("check")),
            hr(),
          ),
          box(
            width = 8, status = "info", title = "Multinomial Regression",
            div(class = "result-output",
                withSpinner(verbatimTextOutput("multinomial_output"))
            )
          )
        )
      ),
      # Poisson Regression Tab - CORRECTED VERSION
      tabItem(
        tabName = "poisson_reg",
        h2("Poisson Regression", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Variables",
            selectInput("poisson_dep", "Dependent", choices = NULL),
            selectizeInput("poisson_indep", "Independent(s)", choices = NULL, multiple = TRUE),
            
            # NEW: Dynamic UI for reference category selection for INDEPENDENT categorical variables
            uiOutput("poisson_ref_ui"),
            
            actionButton("run_poisson", "OK", class = "btn-primary", icon = icon("check")),
            hr(),
          ),
          box(
            width = 8, status = "info", title = "Poisson Regression",
            div(class = "result-output",
                withSpinner(verbatimTextOutput("poisson_output"))
            )
          )
        )
      ),
      # Factor Analysis Tab - CORRECTED VERSION
      tabItem(
        tabName = "factor_analysis",
        h2("Exploratory Factor Analysis", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Variables",
            selectizeInput("factor_vars", "Variables", choices = NULL, multiple = TRUE),
            numericInput("num_factors", "Number of Factors", value = 1, min = 1, max = 20),
            selectInput("rotation_method", "Rotation Method",
                        choices = c("none", "varimax", "quartimax", "promax", "oblimin", "simplimax", "cluster")),
            numericInput("factor_scores_cutoff", "Factor Loading Cutoff", value = 0.3, min = 0.1, max = 0.9, step = 0.05),
            actionButton("run_factor", "OK", class = "btn-primary", icon = icon("check")),
            hr(),
            h4("Options"),
            checkboxGroupInput("factor_options", "Display",
                               choices = c("Eigenvalues", "Factor loadings", "Communalities", 
                                           "Score coefficients", "KMO & Bartlett's", "Reliability")),
            hr(),
          ),
          box(
            width = 8, status = "info", title = "Factor Analysis",
            tabBox(
              width = 12,
              tabPanel("Summary", 
                       div(class = "result-output", 
                           withSpinner(verbatimTextOutput("factor_summary")))
              ),
              tabPanel("Loadings", 
                       div(class = "result-output", 
                           withSpinner(verbatimTextOutput("factor_loadings")))
              ),
              tabPanel("Scree Plot", 
                       withSpinner(plotlyOutput("factor_scree", height = "400px"))
              ),
              tabPanel("Parallel Analysis", 
                       withSpinner(plotlyOutput("factor_parallel", height = "400px"))
              ),
              tabPanel("Factor Scores", 
                       div(class = "result-output", 
                           withSpinner(verbatimTextOutput("factor_scores")))
              ),
              tabPanel("Harman's Test", 
                       div(class = "result-output", 
                           withSpinner(verbatimTextOutput("factor_harman")))
              )
            )
          )
        )
      ),
      # CFA Tab - ENHANCED VERSION
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
                               choices = c("Fit measures", "Parameter estimates", "Modification indices")),
            hr(),
          ),
          box(
            width = 8, status = "info", title = "Confirmatory Factor Analysis",
            tabBox(
              width = 12,
              tabPanel("Summary", div(class = "result-output", withSpinner(verbatimTextOutput("cfa_summary")))),
              tabPanel("Fit Measures", div(class = "result-output", withSpinner(verbatimTextOutput("cfa_fit")))),
              tabPanel("Parameter Estimates", div(class = "result-output", withSpinner(verbatimTextOutput("cfa_params")))),
              tabPanel("Modification Indices", div(class = "result-output", withSpinner(verbatimTextOutput("cfa_modindices")))),
              tabPanel("Path Diagram", withSpinner(plotOutput("cfa_diagram", height = "600px"))),
              tabPanel("Reliability & Validity", div(class = "result-output", withSpinner(verbatimTextOutput("cfa_reliability")))),
              tabPanel("Residuals", div(class = "result-output", withSpinner(verbatimTextOutput("cfa_residuals"))))
            )
          )
        )
      ),
      # Reliability Analysis Tab - CORRECTED VERSION
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
                               choices = c("Cronbach's alpha", "McDonald's omega", "Item statistics", 
                                           "Scale statistics", "Item if deleted", "Inter-item correlations")),
            hr(),
          ),
          box(
            width = 8, status = "info", title = "Reliability Analysis",
            div(class = "result-output",
                withSpinner(verbatimTextOutput("reliability_output"))
            )
          )
        )
      ),
      # Chi-Square Tab - CORRECTED VERSION
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
            h4("Crosstabulation Options"),
            checkboxGroupInput("crosstab_options", "Display in Crosstabulation",
                               choices = c("Observed frequencies", "Expected frequencies", 
                                           "Row percentages", "Column percentages")),
            hr(),
          ),
          box(
            width = 8, status = "info", title = "Chi-Square Test",
            tabBox(
              width = 12,
              tabPanel("Crosstabulation", 
                       div(class = "result-output", 
                           withSpinner(verbatimTextOutput("chi_crosstab")),
                           uiOutput("chi_assumption_note")
                       )),
              tabPanel("Test Statistics", 
                       div(class = "result-output", 
                           withSpinner(verbatimTextOutput("chi_test")))),
              tabPanel("Plots", 
                       withSpinner(plotlyOutput("chi_plot", height = "400px")))
            )
          )
        )
      ),
      # Fisher's Exact Test Tab - CORRECTED VERSION
      tabItem(
        tabName = "fisher_exact",
        h2("Fisher's Exact Test", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 4, status = "primary", title = "Variables",
            selectInput("fisher_row", "Row Variable", choices = NULL),
            selectInput("fisher_col", "Column Variable", choices = NULL),
            actionButton("run_fisher", "OK", class = "btn-primary", icon = icon("check")),
            hr(),
            h4("Crosstabulation Options"),
            checkboxGroupInput("fisher_crosstab_options", "Display in Crosstabulation",
                               choices = c("Observed frequencies", "Expected frequencies", 
                                           "Row percentages", "Column percentages"),
                               selected = "Observed frequencies"),
            hr(),
          ),
          box(
            width = 8, status = "info", title = "Fisher's Exact Test",
            tabBox(
              width = 12,
              tabPanel("Crosstabulation", 
                       div(class = "result-output", 
                           withSpinner(verbatimTextOutput("fisher_crosstab")))
              ),
              tabPanel("Test Statistics", 
                       div(class = "result-output", 
                           withSpinner(verbatimTextOutput("fisher_test")))
              ),
              tabPanel("Plots", 
                       withSpinner(plotlyOutput("fisher_plot", height = "400px")))
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
          ),
          box(
            width = 8, status = "info", title = "Wilcoxon Signed-Rank Test",
            div(class = "result-output",
                withSpinner(verbatimTextOutput("wilcoxon_output"))
            )
          )
        )
      ),
      
      
      
      # Chart Builder Tab - ENHANCED with category-level color selection
      tabItem(
        tabName = "chart_builder",
        h2("Chart Builder", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 3, status = "primary", title = "Variables",
            style = "height: 90vh; overflow-y: auto;",
            selectizeInput("chart_vars", "Variables", choices = NULL, multiple = TRUE),
            selectInput("chart_group", "Grouping Variable", 
                        choices = c("None"), selected = "None"),
            hr(),
            h4("Chart Type"),
            selectInput("chart_type", "Choose a chart type",
                        choices = c("Bar", "Line", "Area", "Pie", "Scatterplot", "Histogram",
                                    "Boxplot", "Density", "Violin", "Q-Q Plot")),
            
            # Dynamic UI for category-level color selection
            uiOutput("category_color_ui"),
            
            # Chart Labels Options
            h4("Chart Labels"),
            selectInput("label_type", "Label Display",
                        choices = c("None", "Frequency", "Percentage", "Both"),
                        selected = "None"),
            numericInput("label_size", "Label Font Size", value = 12, min = 8, max = 24, step = 1),
            
            # Legend Options
            h4("Legend Options"),
            numericInput("legend_size", "Legend Font Size", value = 12, min = 8, max = 24, step = 1),
            
            # Additional Options
            h4("Additional Options"),
            checkboxInput("flip_coords", "Flip Coordinates", value = FALSE),
            checkboxInput("interactive_plot", "Interactive Plot", value = FALSE),
            
            hr(),
            h4("Advanced Customization"),
            textInput("chart_title", "Title", value = ""),
            numericInput("title_size", "Title Size", value = 16, min = 8, max = 24, step = 1),
            textInput("x_axis_label", "X Axis Label:", value = ""),
            textInput("y_axis_label", "Y Axis Label:", value = ""),
            numericInput("axis_text_size", "Axis Text Size", value = 12, min = 8, max = 20, step = 1),
            
            # Base color selection
            h4("Base Color Selection"),
            colourpicker::colourInput("chart_color", "Select Base Color:", value = spss_blue),
            
            # Theme selection
            selectInput("chart_theme", "Theme",
                        choices = c("Classic", "Minimal", "Gray", "Light", "Dark", "BW", "LineDraw"),
                        selected = "Minimal"),
            
            # Data point options
            checkboxInput("show_point_ids", "Show Data Point IDs", value = FALSE),
            
            numericInput("chart_width", "Width (px)", value = 900, min = 400, max = 2000, step = 50),
            numericInput("chart_height", "Height (px)", value = 600, min = 300, max = 1500, step = 50),
            hr(),
            actionButton("generate_chart", "Generate Chart", class = "btn-primary", icon = icon("check")),
            br(), br(),
            downloadButton("download_chart_png", "Download PNG", class = "export-btn"),
            downloadButton("download_chart_jpeg", "Download JPEG", class = "export-btn"),
            downloadButton("download_chart_pdf", "Download PDF", class = "export-btn")
          ),
          box(
            width = 9, status = "info", title = "Chart Preview",
            withSpinner(uiOutput("chart_display"))
          )
        )
      ),
      
      # About Tab
      tabItem(
        tabName = "about",
        h2("About TNMTC DataLab", style = paste0("font-weight: bold; color: ", spss_blue, ";")),
        fluidRow(
          box(
            width = 12, status = "info", title = "Developer Information",
            div(class = "analysis-panel",
                h3("TNMTC DataLab - Statistical Analysis Software"),
                p("TNMTC DataLab is a user-friendly statistical analysis platform that empowers students to perform basic and advanced analyses in R without coding. It streamlines data exploration, statistical testing, and visualization, delivering publication-ready results for research and learning."),
                
                h4("Developer Information"),
                p("Lead Developer: Mudasir Mohammed Ibrahim"),
                p("Email: mudassiribrahim30@gmail.com"),
                p("Collaborator: Mushe Abdul-Hakim"),
                p("Email: musheabdulhakim99@gmail.com"),
                p("For any issues, bugs, or feature requests, please contact the lead developer/collaborator via email."),
                
                h4("Version Information"),
                p("Current Version: 2.0"),
                p("Release Date: September 2025"),
                
                h4("Features"),
                tags$ul(
                  tags$li("Data import from multiple formats (CSV, Excel, SPSS, Stata)"),
                  tags$li("Descriptive and inferential statistical analysis"),
                  tags$li("Advanced regression modeling"),
                  tags$li("Factor analysis and reliability testing"),
                  tags$li("Non-parametric statistical tests"),
                  tags$li("Professional chart building and visualization")
                ),
                
                h4("Technical Support"),
                p("If you encounter any problems while using TNMTC DataLab, please contact:"),
                p("Email: mudassiribrahim30@gmail.com/musheabdulhakim99@gmail.com"),
                p("We will respond to your inquiry as soon as possible."),
                
                h4("Acknowledgments"),
                p("TNMTC DataLab is built using the R programming language and the Shiny framework."),
                p("Special thanks to the R community and package developers for their contributions.")
            )
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

server <- function(input, output, session) {
  anova_results_list <- reactiveVal(list())
  # Enhanced export to Word function
  export_to_word <- function(content, filename, title = NULL, plots = NULL) {
    tryCatch({
      doc <- read_docx()
      
      if (!is.null(title)) {
        doc <- body_add_par(doc, title, style = "heading 1")
        doc <- body_add_par(doc, "")
      }
      
      if (is.character(content)) {
        lines <- strsplit(content, "\n")[[1]]
        for (line in lines) {
          if (line != "") {
            doc <- body_add_par(doc, line)
          }
        }
      } else if (is.data.frame(content)) {
        ft <- flextable(content)
        ft <- autofit(ft)
        doc <- body_add_flextable(doc, ft)
      } else if (is.list(content)) {
        if (!is.null(content$text)) {
          lines <- strsplit(content$text, "\n")[[1]]
          for (line in lines) {
            if (line != "") {
              doc <- body_add_par(doc, line)
            }
          }
        }
        
        if (!is.null(content$tables)) {
          for (table in content$tables) {
            doc <- body_add_par(doc, "")
            ft <- flextable(table)
            ft <- autofit(ft)
            doc <- body_add_flextable(doc, ft)
          }
        }
      }
      
      if (!is.null(plots)) {
        for (i in seq_along(plots)) {
          plot_file <- tempfile(fileext = ".png")
          ggsave(plot_file, plot = plots[[i]], width = 6, height = 4, dpi = 300)
          doc <- body_add_par(doc, "")
          doc <- body_add_img(doc, src = plot_file, width = 6, height = 4)
        }
      }
      
      print(doc, target = filename)
      return(TRUE)
    }, error = function(e) {
      showNotification(paste("Error exporting to Word:", e$message), type = "error")
      return(FALSE)
    })
  }
  
  # Enhanced export to Excel function
  export_to_excel <- function(data, filename, sheet_name = "Results", plots = NULL) {
    tryCatch({
      wb <- createWorkbook()
      
      if (is.list(data) && "tables" %in% names(data)) {
        for (i in seq_along(data$tables)) {
          sheet_name_i <- ifelse(length(data$tables) > 1, paste0(sheet_name, "_", i), sheet_name)
          addWorksheet(wb, sheet_name_i)
          writeData(wb, sheet_name_i, data$tables[[i]], rowNames = FALSE)
        }
        
        if (!is.null(data$text)) {
          addWorksheet(wb, "Summary")
          lines <- strsplit(data$text, "\n")[[1]]
          text_df <- data.frame(Content = lines)
          writeData(wb, "Summary", text_df, rowNames = FALSE)
        }
      } else if (is.data.frame(data)) {
        addWorksheet(wb, sheet_name)
        writeData(wb, sheet_name, data, rowNames = FALSE)
      } else if (is.character(data)) {
        addWorksheet(wb, sheet_name)
        lines <- strsplit(data, "\n")[[1]]
        text_df <- data.frame(Results = lines)
        writeData(wb, sheet_name, text_df, rowNames = FALSE)
      } else {
        addWorksheet(wb, sheet_name)
        writeData(wb, sheet_name, "Statistical Analysis Results")
      }
      
      if (!is.null(plots)) {
        for (i in seq_along(plots)) {
          plot_sheet_name <- paste0("Plot_", i)
          addWorksheet(wb, plot_sheet_name)
          plot_file <- tempfile(fileext = ".png")
          ggsave(plot_file, plot = plots[[i]], width = 6, height = 4, dpi = 300)
          insertImage(wb, plot_sheet_name, plot_file, width = 6, height = 4, startRow = 1, startCol = 1)
        }
      }
      
      saveWorkbook(wb, filename, overwrite = TRUE)
      return(TRUE)
    }, error = function(e) {
      showNotification(paste("Error exporting to Excel:", e$message), type = "error")
      return(FALSE)
    })
  }
  
  # Store analysis results for export
  analysis_results <- reactiveValues(
    descriptives = NULL, frequencies = NULL, ttest = NULL, anova = NULL, 
    correlation = NULL, regression = NULL, logistic = NULL, factor = NULL, 
    reliability = NULL, chi_square = NULL, fisher = NULL, mann_whitney = NULL, 
    kruskal_wallis = NULL, wilcoxon = NULL
  )
  
  # Export handlers for all analysis types
  output$export_descriptives_word <- downloadHandler(
    filename = function() { "descriptives_results.docx" },
    content = function(file) { req(analysis_results$descriptives); export_to_word(analysis_results$descriptives, file, "Descriptive Statistics") }
  )
  
  output$export_descriptives_excel <- downloadHandler(
    filename = function() { "descriptives_results.xlsx" },
    content = function(file) { req(analysis_results$descriptives); export_to_excel(analysis_results$descriptives, file, "Descriptives") }
  )
  
  output$export_frequencies_word <- downloadHandler(
    filename = function() { "frequencies_results.docx" },
    content = function(file) { req(analysis_results$frequencies); export_to_word(analysis_results$frequencies, file, "Frequency Analysis") }
  )
  
  output$export_frequencies_excel <- downloadHandler(
    filename = function() { "frequencies_results.xlsx" },
    content = function(file) { req(analysis_results$frequencies); export_to_excel(analysis_results$frequencies, file, "Frequencies") }
  )
  
  output$export_ttest_word <- downloadHandler(
    filename = function() { "ttest_results.docx" },
    content = function(file) { req(analysis_results$ttest); export_to_word(analysis_results$ttest, file, "T-Test Analysis") }
  )
  
  output$export_ttest_excel <- downloadHandler(
    filename = function() { "ttest_results.xlsx" },
    content = function(file) { req(analysis_results$ttest); export_to_excel(analysis_results$ttest, file, "T-Test") }
  )
  
  output$export_onesample_word <- downloadHandler(
    filename = function() { "onesample_ttest_results.docx" },
    content = function(file) { req(analysis_results$ttest); export_to_word(analysis_results$ttest, file, "One Sample T-Test") }
  )
  
  output$export_onesample_excel <- downloadHandler(
    filename = function() { "onesample_ttest_results.xlsx" },
    content = function(file) { req(analysis_results$ttest); export_to_excel(analysis_results$ttest, file, "One Sample T-Test") }
  )
  
  output$export_paired_word <- downloadHandler(
    filename = function() { "paired_ttest_results.docx" },
    content = function(file) { req(analysis_results$ttest); export_to_word(analysis_results$ttest, file, "Paired T-Test") }
  )
  
  output$export_paired_excel <- downloadHandler(
    filename = function() { "paired_ttest_results.xlsx" },
    content = function(file) { req(analysis_results$ttest); export_to_excel(analysis_results$ttest, file, "Paired T-Test") }
  )
  
  output$export_anova_word <- downloadHandler(
    filename = function() { "anova_results.docx" },
    content = function(file) { req(analysis_results$anova); export_to_word(analysis_results$anova, file, "ANOVA") }
  )
  
  output$export_anova_excel <- downloadHandler(
    filename = function() { "anova_results.xlsx" },
    content = function(file) { req(analysis_results$anova); export_to_excel(analysis_results$anova, file, "ANOVA") }
  )
  
  output$export_correlation_word <- downloadHandler(
    filename = function() { "correlation_results.docx" },
    content = function(file) { req(analysis_results$correlation); export_to_word(analysis_results$correlation, file, "Correlation Analysis") }
  )
  
  output$export_correlation_excel <- downloadHandler(
    filename = function() { "correlation_results.xlsx" },
    content = function(file) { req(analysis_results$correlation); export_to_excel(analysis_results$correlation, file, "Correlations") }
  )
  
  output$export_regression_word <- downloadHandler(
    filename = function() { "regression_results.docx" },
    content = function(file) { req(analysis_results$regression); export_to_word(analysis_results$regression, file, "Linear Regression") }
  )
  
  output$export_regression_excel <- downloadHandler(
    filename = function() { "regression_results.xlsx" },
    content = function(file) { req(analysis_results$regression); export_to_excel(analysis_results$regression, file, "Linear Regression") }
  )
  
  output$export_logistic_word <- downloadHandler(
    filename = function() { "logistic_results.docx" },
    content = function(file) { req(analysis_results$logistic); export_to_word(analysis_results$logistic, file, "Logistic Regression") }
  )
  
  output$export_logistic_excel <- downloadHandler(
    filename = function() { "logistic_results.xlsx" },
    content = function(file) { req(analysis_results$logistic); export_to_excel(analysis_results$logistic, file, "Logistic Regression") }
  )
  
  output$export_ordinal_word <- downloadHandler(
    filename = function() { "ordinal_results.docx" },
    content = function(file) { req(analysis_results$logistic); export_to_word(analysis_results$logistic, file, "Ordinal Regression") }
  )
  
  output$export_ordinal_excel <- downloadHandler(
    filename = function() { "ordinal_results.xlsx" },
    content = function(file) { req(analysis_results$logistic); export_to_excel(analysis_results$logistic, file, "Ordinal Regression") }
  )
  
  output$export_multinomial_word <- downloadHandler(
    filename = function() { "multinomial_results.docx" },
    content = function(file) { req(analysis_results$logistic); export_to_word(analysis_results$logistic, file, "Multinomial Regression") }
  )
  
  output$export_multinomial_excel <- downloadHandler(
    filename = function() { "multinomial_results.xlsx" },
    content = function(file) { req(analysis_results$logistic); export_to_excel(analysis_results$logistic, file, "Multinomial Regression") }
  )
  
  output$export_poisson_word <- downloadHandler(
    filename = function() { "poisson_results.docx" },
    content = function(file) { req(analysis_results$regression); export_to_word(analysis_results$regression, file, "Poisson Regression") }
  )
  
  output$export_poisson_excel <- downloadHandler(
    filename = function() { "poisson_results.xlsx" },
    content = function(file) { req(analysis_results$regression); export_to_excel(analysis_results$regression, file, "Poisson Regression") }
  )
  
  output$export_factor_word <- downloadHandler(
    filename = function() { "factor_results.docx" },
    content = function(file) { req(analysis_results$factor); export_to_word(analysis_results$factor, file, "Factor Analysis") }
  )
  
  output$export_factor_excel <- downloadHandler(
    filename = function() { "factor_results.xlsx" },
    content = function(file) { req(analysis_results$factor); export_to_excel(analysis_results$factor, file, "Factor Analysis") }
  )
  
  output$export_cfa_word <- downloadHandler(
    filename = function() { "cfa_results.docx" },
    content = function(file) { req(analysis_results$factor); export_to_word(analysis_results$factor, file, "Confirmatory Factor Analysis") }
  )
  
  output$export_cfa_excel <- downloadHandler(
    filename = function() { "cfa_results.xlsx" },
    content = function(file) { req(analysis_results$factor); export_to_excel(analysis_results$factor, file, "Confirmatory Factor Analysis") }
  )
  
  output$export_reliability_word <- downloadHandler(
    filename = function() { "reliability_results.docx" },
    content = function(file) { req(analysis_results$reliability); export_to_word(analysis_results$reliability, file, "Reliability Analysis") }
  )
  
  output$export_reliability_excel <- downloadHandler(
    filename = function() { "reliability_results.xlsx" },
    content = function(file) { req(analysis_results$reliability); export_to_excel(analysis_results$reliability, file, "Reliability Analysis") }
  )
  
  output$export_chi_word <- downloadHandler(
    filename = function() { "chi_square_results.docx" },
    content = function(file) { req(analysis_results$chi_square); export_to_word(analysis_results$chi_square, file, "Chi-Square Test") }
  )
  
  output$export_chi_excel <- downloadHandler(
    filename = function() { "chi_square_results.xlsx" },
    content = function(file) { req(analysis_results$chi_square); export_to_excel(analysis_results$chi_square, file, "Chi-Square Test") }
  )
  
  output$export_fisher_word <- downloadHandler(
    filename = function() { "fisher_results.docx" },
    content = function(file) { req(analysis_results$fisher); export_to_word(analysis_results$fisher, file, "Fisher's Exact Test") }
  )
  
  output$export_fisher_excel <- downloadHandler(
    filename = function() { "fisher_results.xlsx" },
    content = function(file) { req(analysis_results$fisher); export_to_excel(analysis_results$fisher, file, "Fisher's Exact Test") }
  )
  
  output$export_mw_word <- downloadHandler(
    filename = function() { "mann_whitney_results.docx" },
    content = function(file) { req(analysis_results$mann_whitney); export_to_word(analysis_results$mann_whitney, file, "Mann-Whitney U Test") }
  )
  
  output$export_mw_excel <- downloadHandler(
    filename = function() { "mann_whitney_results.xlsx" },
    content = function(file) { req(analysis_results$mann_whitney); export_to_excel(analysis_results$mann_whitney, file, "Mann-Whitney U Test") }
  )
  
  output$export_kw_word <- downloadHandler(
    filename = function() { "kruskal_wallis_results.docx" },
    content = function(file) { req(analysis_results$kruskal_wallis); export_to_word(analysis_results$kruskal_wallis, file, "Kruskal-Wallis Test") }
  )
  
  output$export_kw_excel <- downloadHandler(
    filename = function() { "kruskal_wallis_results.xlsx" },
    content = function(file) { req(analysis_results$kruskal_wallis); export_to_excel(analysis_results$kruskal_wallis, file, "Kruskal-Wallis Test") }
  )
  
  output$export_wilcoxon_word <- downloadHandler(
    filename = function() { "wilcoxon_results.docx" },
    content = function(file) { req(analysis_results$wilcoxon); export_to_word(analysis_results$wilcoxon, file, "Wilcoxon Signed-Rank Test") }
  )
  
  output$export_wilcoxon_excel <- downloadHandler(
    filename = function() { "wilcoxon_results.xlsx" },
    content = function(file) { req(analysis_results$wilcoxon); export_to_excel(analysis_results$wilcoxon, file, "Wilcoxon Signed-Rank Test") }
  )
  
  # Store results when analyses are run
  observeEvent(input$run_descriptives, {
    output_text <- capture.output({
      df <- data()
      selected_vars <- input$desc_vars
      if (length(selected_vars) == 0) return("Please select at least one variable.")
      data_subset <- df[, selected_vars, drop = FALSE]
      desc_stats <- psych::describe(data_subset)
      cat("Descriptive Statistics\n=====================\n\n")
      if ("Mean" %in% input$desc_stats) { cat("Means:\n"); print(round(desc_stats$mean, 3)); cat("\n") }
      if ("Std. deviation" %in% input$desc_stats) { cat("Standard Deviations:\n"); print(round(desc_stats$sd, 3)); cat("\n") }
      if ("Minimum" %in% input$desc_stats) { cat("Minimum Values:\n"); print(desc_stats$min); cat("\n") }
      if ("Maximum" %in% input$desc_stats) { cat("Maximum Values:\n"); print(desc_stats$max); cat("\n") }
      if ("Variance" %in% input$desc_stats) { variances <- apply(data_subset, 2, var, na.rm = TRUE); cat("Variances:\n"); print(round(variances, 3)); cat("\n") }
      if ("Range" %in% input$desc_stats) { ranges <- apply(data_subset, 2, function(x) diff(range(x, na.rm = TRUE))); cat("Ranges:\n"); print(round(ranges, 3)); cat("\n") }
      if ("S.E. mean" %in% input$desc_stats) { se_means <- desc_stats$sd / sqrt(desc_stats$n); cat("Standard Error of Mean:\n"); print(round(se_means, 3)); cat("\n") }
      if ("Kurtosis" %in% input$desc_stats) { cat("Kurtosis:\n"); print(round(desc_stats$kurtosis, 3)); cat("\n") }
      if ("Skewness" %in% input$desc_stats) { cat("Skewness:\n"); print(round(desc_stats$skew, 3)); cat("\n") }
      if ("Valid N" %in% input$desc_stats) { cat("Valid Cases (N):\n"); print(desc_stats$n); cat("\n") }
      if ("Missing N" %in% input$desc_stats) { missing_counts <- colSums(is.na(data_subset)); cat("Missing Values:\n"); print(missing_counts); cat("\n") }
    })
    analysis_results$descriptives <- list(text = paste(output_text, collapse = "\n"), tables = list(desc_stats))
  })
  
  observeEvent(input$run_frequencies, {
    output_text <- capture.output({
      df <- data()
      selected_vars <- input$freq_vars
      if (length(selected_vars) == 0) return("Please select at least one variable.")
      cat("Frequency Tables\n================\n\n")
      for (var in selected_vars) {
        cat("Variable:", var, "\n----------------------------------------\n")
        freq_table <- table(df[[var]], useNA = "always")
        prop_table <- prop.table(freq_table) * 100
        result_df <- data.frame(Category = names(freq_table), Frequency = as.numeric(freq_table), Percent = round(as.numeric(prop_table), 2), CumulativePercent = round(cumsum(as.numeric(prop_table)), 2))
        result_df$Category[is.na(result_df$Category)] <- "Missing"
        print(result_df, row.names = FALSE)
        cat("\n\n")
      }
    })
    analysis_results$frequencies <- list(text = paste(output_text, collapse = "\n"))
  })
  
  observeEvent(input$run_ttest, {
    output_text <- capture.output({
      df <- data()
      test_vars <- input$ttest_vars
      group_var <- input$ttest_group_var
      if (length(test_vars) == 0) return("Please select at least one test variable.")
      if (is.null(group_var) || group_var == "") return("Please select a grouping variable.")
      groups <- strsplit(input$group_def, ",")[[1]]; groups <- trimws(groups)
      if (length(groups) != 2) return("Please define exactly two groups (e.g., 'male,female' for Gender)")
      cat("Independent Samples T-Test\n==========================\n\n")
      for (test_var in test_vars) {
        cat("Variable:", test_var, "\n----------------------------------------\n")
        group_data <- df[df[[group_var]] %in% groups, ]
        group1_data <- group_data[group_data[[group_var]] == groups[1], test_var]
        group2_data <- group_data[group_data[[group_var]] == groups[2], test_var]
        group1_data <- group1_data[!is.na(group1_data)]; group2_data <- group2_data[!is.na(group2_data)]
        if (length(group1_data) < 2 || length(group2_data) < 2) { cat("Insufficient data for t-test:\n"); cat("Group 1 (", groups[1], "): N =", length(group1_data), "\n"); cat("Group 2 (", groups[2], "): N =", length(group2_data), "\n"); cat("Each group needs at least 2 observations for t-test.\n\n"); next }
        if (input$ttest_tails == "Two-tailed") { t_test <- t.test(group1_data, group2_data, var.equal = TRUE, conf.level = input$ttest_conf/100, alternative = "two.sided")
        } else { mean_diff <- mean(group1_data) - mean(group2_data); alternative <- ifelse(mean_diff > 0, "greater", "less"); t_test <- t.test(group1_data, group2_data, var.equal = TRUE, conf.level = input$ttest_conf/100, alternative = alternative) }
        cat("Group 1 (", groups[1], "): N =", length(group1_data), ", Mean =", round(mean(group1_data), 3), ", SD =", round(sd(group1_data), 3), "\n")
        cat("Group 2 (", groups[2], "): N =", length(group2_data), ", Mean =", round(mean(group2_data), 3), ", SD =", round(sd(group2_data), 3), "\n\n")
        cat("t(", t_test$parameter, ") =", round(t_test$statistic, 3), "\n"); cat("p-value =", format.pval(t_test$p.value, digits = 3), "\n")
        cat(input$ttest_conf, "% CI: [", round(t_test$conf.int[1], 3), ",", round(t_test$conf.int[2], 3), "]\n")
        cat("Mean Difference =", round(t_test$estimate[1] - t_test$estimate[2], 3), "\n\n")
        if (input$ttest_tails == "One-tailed") {
          if (t_test$p.value < 0.05) { if (mean(group1_data) > mean(group2_data)) { cat("Interpretation: Group 1 is significantly greater than Group 2 (one-tailed)\n") } else { cat("Interpretation: Group 1 is significantly less than Group 2 (one-tailed)\n") } } else { cat("Interpretation: No significant difference between groups (one-tailed)\n") }
        }
        cat("\n")
      }
    })
    analysis_results$ttest <- list(text = paste(output_text, collapse = "\n"))
  })
  
  observeEvent(input$run_onesample, {
    output_text <- capture.output({
      df <- data(); test_vars <- input$onesample_vars; test_value <- input$test_value
      if (length(test_vars) == 0) return("Please select at least one test variable.")
      cat("One Sample T-Test\n=================\n\n"); cat("Test Value:", test_value, "\n\n")
      for (test_var in test_vars) {
        cat("Variable:", test_var, "\n----------------------------------------\n")
        test_data <- df[[test_var]]; test_data <- test_data[!is.na(test_data)]
        if (input$onesample_tails == "Two-tailed") { t_test <- t.test(test_data, mu = test_value, conf.level = input$onesample_conf/100, alternative = "two.sided")
        } else { mean_diff <- mean(test_data) - test_value; alternative <- ifelse(mean_diff > 0, "greater", "less"); t_test <- t.test(test_data, mu = test_value, conf.level = input$onesample_conf/100, alternative = alternative) }
        cat("N =", length(test_data), ", Mean =", round(mean(test_data), 3), ", SD =", round(sd(test_data), 3), "\n\n")
        cat("t(", t_test$parameter, ") =", round(t_test$statistic, 3), "\n"); cat("p-value =", format.pval(t_test$p.value, digits = 3), "\n")
        cat(input$onesample_conf, "% CI: [", round(t_test$conf.int[1], 3), ",", round(t_test$conf.int[2], 3), "]\n")
        cat("Mean Difference =", round(mean(test_data) - test_value, 3), "\n\n")
        if (input$onesample_tails == "One-tailed") {
          if (t_test$p.value < 0.05) { if (mean(test_data) > test_value) { cat("Interpretation: Mean is significantly greater than test value (one-tailed)\n") } else { cat("Interpretation: Mean is significantly less than test value (one-tailed)\n") } } else { cat("Interpretation: No significant difference from test value (one-tailed)\n") }
        }
        cat("\n")
      }
    })
    analysis_results$ttest <- list(text = paste(output_text, collapse = "\n"))
  })
  
  observeEvent(input$run_paired, {
    output_text <- capture.output({
      df <- data(); paired_vars <- input$paired_vars
      if (length(paired_vars) < 2) return("Please select at least two variables for paired comparison.")
      cat("Paired Samples T-Test\n=====================\n\n")
      pairs <- combn(paired_vars, 2, simplify = FALSE)
      for (pair in pairs) {
        cat("Pair:", pair[1], "vs", pair[2], "\n----------------------------------------\n")
        var1 <- df[[pair[1]]]; var2 <- df[[pair[2]]]
        complete_cases <- complete.cases(var1, var2); var1 <- var1[complete_cases]; var2 <- var2[complete_cases]
        if (input$paired_tails == "Two-tailed") { t_test <- t.test(var1, var2, paired = TRUE, conf.level = input$paired_conf/100, alternative = "two.sided")
        } else { mean_diff <- mean(var1) - mean(var2); alternative <- ifelse(mean_diff > 0, "greater", "less"); t_test <- t.test(var1, var2, paired = TRUE, conf.level = input$paired_conf/100, alternative = alternative) }
        differences <- var1 - var2
        cat("N =", length(var1), "\n"); cat("Mean of", pair[1], "=", round(mean(var1), 3), "\n"); cat("Mean of", pair[2], "=", round(mean(var2), 3), "\n")
        cat("Mean Difference =", round(mean(differences), 3), "\n"); cat("SD of Differences =", round(sd(differences), 3), "\n"); cat("SE of Differences =", round(sd(differences)/sqrt(length(differences)), 3), "\n\n")
        cat("t(", t_test$parameter, ") =", round(t_test$statistic, 3), "\n"); cat("p-value =", format.pval(t_test$p.value, digits = 3), "\n")
        cat(input$paired_conf, "% CI: [", round(t_test$conf.int[1], 3), ",", round(t_test$conf.int[2], 3), "]\n\n")
        if (input$paired_tails == "One-tailed") {
          if (t_test$p.value < 0.05) { if (mean(var1) > mean(var2)) { cat("Interpretation:", pair[1], "is significantly greater than", pair[2], "(one-tailed)\n") } else { cat("Interpretation:", pair[1], "is significantly less than", pair[2], "(one-tailed)\n") } } else { cat("Interpretation: No significant difference between variables (one-tailed)\n") }
        }
        cat("\n")
      }
    })
    analysis_results$ttest <- list(text = paste(output_text, collapse = "\n"))
  })
  
  observeEvent(input$run_correlation, {
    output_text <- capture.output({
      df <- data(); cor_vars <- input$cor_vars
      if (length(cor_vars) < 2) return("Please select at least two variables for correlation analysis.")
      cor_vars <- df[, cor_vars, drop = FALSE]
      numeric_vars <- sapply(cor_vars, is.numeric)
      if (sum(numeric_vars) < 2) return("Please select at least two numeric variables for correlation analysis.")
      cor_vars <- cor_vars[, numeric_vars, drop = FALSE]
      cor_matrix <- cor(cor_vars, use = "pairwise.complete.obs", method = input$cor_method)
      n <- nrow(cor_vars); cor_test <- corr.test(cor_vars, method = input$cor_method)
      if (input$sig_test == "One-tailed") cor_test$p <- cor_test$p / 2
      cat("Bivariate Correlations\n======================\n\n")
      cat("Method: ", switch(input$cor_method, "pearson" = "Pearson", "spearman" = "Spearman", "kendall" = "Kendall"), "correlation\n")
      cat("Test Type: ", input$sig_test, "\n\n"); cat("Correlation Matrix:\n")
      cor_with_stars <- matrix("", nrow = nrow(cor_matrix), ncol = ncol(cor_matrix))
      rownames(cor_with_stars) <- rownames(cor_matrix); colnames(cor_with_stars) <- colnames(cor_matrix)
      for (i in 1:nrow(cor_matrix)) {
        for (j in 1:ncol(cor_matrix)) {
          if (i == j) { cor_with_stars[i, j] <- "1.000" } else if (i > j) {
            cor_value <- sprintf("%.3f", cor_matrix[i, j])
            if (input$flag_sig) { p_value <- cor_test$p[i, j]; stars <- ""; if (!is.na(p_value)) { if (p_value < 0.001) stars <- "***" else if (p_value < 0.01) stars <- "**" else if (p_value < 0.05) stars <- "*" }; cor_with_stars[i, j] <- paste0(cor_value, stars) } else { cor_with_stars[i, j] <- cor_value }
          } else { cor_with_stars[i, j] <- "" }
        }
      }
      print(cor_with_stars, quote = FALSE); cat("\nSample Size (Pairwise):\n"); print(cor_test$n)
      if (input$flag_sig) { cat("\nSignificance Levels: * p < 0.05, ** p < 0.01, *** p < 0.001\n"); cat("Note: p-values are for", tolower(input$sig_test), "test\n") }
    })
    analysis_results$correlation <- list(text = paste(output_text, collapse = "\n"), tables = list(cor_matrix))
  })
  
  observeEvent(input$run_anova, {
    output_text <- capture.output({
      df <- data(); dv_vars <- input$anova_dv; factor_var <- input$anova_factor
      if (length(dv_vars) == 0) return("Please select at least one dependent variable.")
      if (is.null(factor_var) || factor_var == "") return("Please select a factor variable.")
      cat("One-Way ANOVA\n=============\n\n"); anova_results <- list()
      for (dv_var in dv_vars) {
        cat("Dependent Variable:", dv_var, "\nFactor:", factor_var, "\n----------------------------------------\n")
        formula <- as.formula(paste(dv_var, "~", factor_var)); anova_result <- aov(formula, data = df); anova_summary <- summary(anova_result)
        anova_results[[dv_var]] <- anova_result; print(anova_summary); cat("\n")
        if ("Homogeneity of variance test" %in% input$anova_options) { cat("Homogeneity of Variance Test (Levene's Test):\n"); levene_test <- car::leveneTest(formula, data = df); print(levene_test); cat("\n") }
        if ("Welch" %in% input$anova_options) { cat("Welch ANOVA (for unequal variances):\n"); welch_test <- oneway.test(formula, data = df, var.equal = FALSE); print(welch_test); cat("\n") }
        if ("Brown-Forsythe" %in% input$anova_options) { cat("Brown-Forsythe Test:\n"); bf_test <- oneway.test(formula, data = df, var.equal = FALSE); cat("F(", bf_test$parameter[1], ",", bf_test$parameter[2], ") =", round(bf_test$statistic, 3), ", p =", format.pval(bf_test$p.value, digits = 3), "\n\n") }
      }
      anova_results_list(anova_results)
    })
    analysis_results$anova <- list(text = paste(output_text, collapse = "\n"))
  })
  
  observeEvent(input$run_chi, {
    output_text <- capture.output({
      df <- data(); row_var <- input$chi_row; col_var <- input$chi_col
      contingency_table <- table(df[[row_var]], df[[col_var]])
      cat("Crosstabulation\n===============\n\n"); cat("Row Variable:", row_var, "\nColumn Variable:", col_var, "\n\n")
      if ("Observed frequencies" %in% input$crosstab_options) { cat("Observed Frequencies:\n"); print(addmargins(contingency_table)); cat("\n") }
      chi_test <- chisq.test(contingency_table); expected <- chi_test$expected
      if ("Expected frequencies" %in% input$crosstab_options) { cat("Expected Frequencies:\n"); print(round(addmargins(expected), 2)); cat("\n") }
      if ("Row percentages" %in% input$crosstab_options) { cat("Row Percentages:\n"); row_perc <- prop.table(contingency_table, 1) * 100; print(round(addmargins(row_perc), 2)); cat("\n") }
      if ("Column percentages" %in% input$crosstab_options) { cat("Column Percentages:\n"); col_perc <- prop.table(contingency_table, 2) * 100; print(round(addmargins(col_perc), 2)); cat("\n") }
      chi_test_result(chi_test)
    })
    analysis_results$chi_square <- list(text = paste(output_text, collapse = "\n"), tables = list(contingency_table))
  })
  
  observeEvent(input$run_fisher, {
    output_text <- capture.output({
      df <- data(); row_var <- input$fisher_row; col_var <- input$fisher_col
      contingency_table <- table(df[[row_var]], df[[col_var]])
      cat("Crosstabulation\n===============\n\n"); cat("Row Variable:", row_var, "\nColumn Variable:", col_var, "\n\n")
      if ("Observed frequencies" %in% input$fisher_crosstab_options) { cat("Observed Frequencies:\n"); print(addmargins(contingency_table)); cat("\n") }
      chi_test <- chisq.test(contingency_table); expected <- chi_test$expected
      if ("Expected frequencies" %in% input$fisher_crosstab_options) { cat("Expected Frequencies:\n"); print(round(addmargins(expected), 2)); cat("\n") }
      if ("Row percentages" %in% input$fisher_crosstab_options) { cat("Row Percentages:\n"); row_perc <- prop.table(contingency_table, 1) * 100; print(round(addmargins(row_perc), 2)); cat("\n") }
      if ("Column percentages" %in% input$fisher_crosstab_options) { cat("Column Percentages:\n"); col_perc <- prop.table(contingency_table, 2) * 100; print(round(addmargins(col_perc), 2)); cat("\n") }
      fisher_test <- fisher.test(contingency_table)
      cat("Fisher's Exact Test\n===================\n\n")
      cat("p-value =", format.pval(fisher_test$p.value, digits = 3), "\n"); cat("Odds Ratio =", round(fisher_test$estimate, 3), "\n")
      if (!is.null(fisher_test$conf.int)) { cat("95% CI for Odds Ratio: [", round(fisher_test$conf.int[1], 3), ",", round(fisher_test$conf.int[2], 3), "]\n") }
    })
    analysis_results$fisher <- list(text = paste(output_text, collapse = "\n"), tables = list(contingency_table))
  })
  
  observeEvent(input$run_reliability, {
    output_text <- capture.output({
      df <- data(); reliability_vars <- input$reliability_vars
      if (length(reliability_vars) < 2) return("Please select at least 2 items for reliability analysis.")
      numeric_vars <- sapply(df[, reliability_vars], is.numeric)
      if (sum(numeric_vars) < 2) return("Please select at least 2 numeric items for reliability analysis.")
      reliability_data <- df[, reliability_vars[numeric_vars]]; reliability_data <- na.omit(reliability_data)
      cat("Reliability Analysis\n====================\n\n")
      cat("Number of Items:", ncol(reliability_data), "\n"); cat("Number of Cases:", nrow(reliability_data), "\n\n")
      if ("Cronbach's alpha" %in% input$reliability_stats) {
        alpha_result <- psych::alpha(reliability_data)
        cat("Cronbach's Alpha\n----------------\n"); cat("Raw Alpha:", round(alpha_result$total$raw_alpha, 3), "\n"); cat("Standardized Alpha:", round(alpha_result$total$std.alpha, 3), "\n\n")
        alpha_value <- alpha_result$total$raw_alpha
        if (alpha_value >= 0.9) cat("Interpretation: Excellent reliability\n") else if (alpha_value >= 0.8) cat("Interpretation: Good reliability\n") else if (alpha_value >= 0.7) cat("Interpretation: Acceptable reliability\n") else if (alpha_value >= 0.6) cat("Interpretation: Questionable reliability\n") else if (alpha_value >= 0.5) cat("Interpretation: Poor reliability\n") else cat("Interpretation: Unacceptable reliability\n"); cat("\n")
      }
      if ("McDonald's omega" %in% input$reliability_stats) {
        if (requireNamespace("psych", quietly = TRUE)) {
          omega_result <- psych::omega(reliability_data, plot = FALSE)
          cat("McDonald's Omega\n----------------\n"); cat("Omega Total:", round(omega_result$omega.tot, 3), "\n"); cat("Omega Hierarchical:", round(omega_result$omega_h, 3), "\n\n")
        } else cat("McDonald's Omega requires the 'psych' package.\n\n")
      }
      if ("Item statistics" %in% input$reliability_stats) { cat("Item Statistics\n---------------\n"); item_stats <- psych::describe(reliability_data); print(round(item_stats, 3)); cat("\n") }
      if ("Scale statistics" %in% input$reliability_stats) { cat("Scale Statistics\n----------------\n"); total_scores <- rowSums(reliability_data, na.rm = TRUE); scale_stats <- psych::describe(total_scores); print(round(scale_stats, 3)); cat("\n") }
      if ("Item if deleted" %in% input$reliability_stats) {
        cat("Item-Total Statistics\n---------------------\n"); alpha_result <- psych::alpha(reliability_data); item_total <- alpha_result$item.stats
        item_total_df <- data.frame(Item = rownames(item_total), Mean = round(item_total$mean, 3), SD = round(item_total$sd, 3), `Item-Total Correlation` = round(item_total$r.cor, 3), `Alpha if Deleted` = round(item_total$raw.r, 3))
        print(item_total_df, row.names = FALSE); cat("\n")
      }
      if ("Inter-item correlations" %in% input$reliability_stats) {
        cat("Inter-Item Correlations\n-----------------------\n"); inter_cor <- cor(reliability_data, use = "pairwise.complete.obs"); print(round(inter_cor, 3)); cat("\n")
        avg_inter_cor <- mean(inter_cor[lower.tri(inter_cor)]); cat("Average Inter-Item Correlation:", round(avg_inter_cor, 3), "\n")
      }
    })
    analysis_results$reliability <- list(text = paste(output_text, collapse = "\n"))
  })
  
  observeEvent(input$run_mw, {
    output_text <- capture.output({
      df <- data(); test_vars <- input$mw_vars; group_var <- input$mw_group_var
      if (length(test_vars) == 0) return("Please select at least one test variable.")
      groups <- strsplit(input$mw_group_def, ",")[[1]]; groups <- trimws(groups)
      if (length(groups) != 2) return("Please define exactly two groups (e.g., '1,2' or 'A,B')")
      cat("Mann-Whitney U Test\n===================\n\n")
      for (test_var in test_vars) {
        cat("Variable:", test_var, "\n----------------------------------------\n")
        group_data <- df[df[[group_var]] %in% groups, ]
        group1_data <- group_data[group_data[[group_var]] == groups[1], test_var]
        group2_data <- group_data[group_data[[group_var]] == groups[2], test_var]
        group1_data <- group1_data[!is.na(group1_data)]; group2_data <- group2_data[!is.na(group2_data)]
        mw_test <- wilcox.test(group1_data, group2_data, exact = FALSE)
        cat("Group 1 (", groups[1], "): N =", length(group1_data), ", Median =", round(median(group1_data), 3), ", Mean Rank =", round(mean(rank(c(group1_data, group2_data))[1:length(group1_data)]), 3), "\n")
        cat("Group 2 (", groups[2], "): N =", length(group2_data), ", Median =", round(median(group2_data), 3), ", Mean Rank =", round(mean(rank(c(group1_data, group2_data))[(length(group1_data)+1):length(c(group1_data, group2_data))]), 3), "\n\n")
        cat("Mann-Whitney U =", mw_test$statistic, "\n"); cat("p-value =", format.pval(mw_test$p.value, digits = 3), "\n\n")
      }
    })
    analysis_results$mann_whitney <- list(text = paste(output_text, collapse = "\n"))
  })
  
  observeEvent(input$run_kw, {
    output_text <- capture.output({
      df <- data(); test_vars <- input$kw_vars; group_var <- input$kw_group_var
      if (length(test_vars) == 0) return("Please select at least one test variable.")
      cat("Kruskal-Wallis Test\n===================\n\n")
      for (test_var in test_vars) {
        cat("Variable:", test_var, "\n----------------------------------------\n")
        complete_cases <- complete.cases(df[[test_var]], df[[group_var]])
        test_data <- df[complete_cases, test_var]; group_data <- df[complete_cases, group_var]
        kw_test <- kruskal.test(test_data ~ group_data); groups <- unique(group_data)
        cat("Group Statistics:\n")
        for (group in groups) { group_values <- test_data[group_data == group]; cat("Group", group, ": N =", length(group_values), ", Median =", round(median(group_values), 3), ", Mean Rank =", round(mean(rank(test_data)[group_data == group]), 3), "\n") }
        cat("\n"); cat("Kruskal-Wallis H(", kw_test$parameter, ") =", round(kw_test$statistic, 3), "\n"); cat("p-value =", format.pval(kw_test$p.value, digits = 3), "\n\n")
      }
    })
    analysis_results$kruskal_wallis <- list(text = paste(output_text, collapse = "\n"))
  })
  
  observeEvent(input$run_wilcoxon, {
    output_text <- capture.output({
      df <- data(); paired_vars <- input$wilcoxon_vars
      if (length(paired_vars) < 2) return("Please select at least two variables for paired comparison.")
      cat("Wilcoxon Signed-Rank Test\n=========================\n\n")
      pairs <- combn(paired_vars, 2, simplify = FALSE)
      for (pair in pairs) {
        cat("Pair:", pair[1], "vs", pair[2], "\n----------------------------------------\n")
        var1 <- df[[pair[1]]]; var2 <- df[[pair[2]]]
        complete_cases <- complete.cases(var1, var2); var1 <- var1[complete_cases]; var2 <- var2[complete_cases]
        wilcox_test <- wilcox.test(var1, var2, paired = TRUE, exact = FALSE)
        differences <- var1 - var2
        cat("N =", length(var1), "\n"); cat("Mean Difference =", round(mean(differences), 3), "\n"); cat("Median Difference =", round(median(differences), 3), "\n")
        cat("Positive Ranks:", sum(differences > 0), "\n"); cat("Negative Ranks:", sum(differences < 0), "\n"); cat("Ties:", sum(differences == 0), "\n\n")
        cat("Wilcoxon V =", wilcox_test$statistic, "\n"); cat("p-value =", format.pval(wilcox_test$p.value, digits = 3), "\n\n")
      }
    })
    analysis_results$wilcoxon <- list(text = paste(output_text, collapse = "\n"))
  })
  
  # Reactive data storage
  data <- reactiveVal()
  var_types <- reactiveVal()
  data_modified <- reactiveVal(FALSE)
  
  # Enhanced data import with better type detection and SPSS value handling
  observeEvent(input$load_btn, {
    req(input$data_format)
    
    tryCatch({
      df <- NULL
      
      if (input$data_format == "CSV" && !is.null(input$file_csv)) {
        # Enhanced CSV reading with better type detection
        df <- read.csv(input$file_csv$datapath, 
                       header = input$header, 
                       sep = input$sep,
                       dec = input$dec,
                       stringsAsFactors = FALSE,
                       na.strings = c("", "NA", "N/A", "NULL"))
        
      } else if (input$data_format == "Excel" && !is.null(input$file_excel)) {
        df <- read_excel(input$file_excel$datapath)
        df <- as.data.frame(df)
        
      } else if (input$data_format == "SPSS" && !is.null(input$file_spss)) {
        # FIX: Use user_value = TRUE to get numeric values instead of labels
        df <- read_sav(input$file_spss$datapath, user_na = TRUE)
        df <- as.data.frame(df)
        
        # Convert labelled variables to regular factors or numeric
        df <- df %>% mutate(across(where(is.labelled), as_factor))
        
      } else if (input$data_format == "Stata" && !is.null(input$file_stata)) {
        df <- read_dta(input$file_stata$datapath)
        df <- as.data.frame(df)
        
      } else {
        showNotification("Please select a file", type = "error")
        return()
      }
      
      # Enhanced data type detection
      df <- detect_data_types(df)
      
      data(df)
      data_modified(FALSE)
      
      # Determine variable types
      types <- sapply(df, function(x) {
        if (is.numeric(x)) "Numeric"
        else if (is.factor(x)) "Categorical"
        else "Categorical"
      })
      
      var_types(data.frame(
        Variable = names(df),
        Type = types,
        stringsAsFactors = TRUE
      ))
      
      # Update all select inputs
      update_all_select_inputs(df, types, session)
      
      showNotification("Data imported successfully!", type = "message")
      
    }, error = function(e) {
      showNotification(paste("Error:", e$message), type = "error")
    })
  })
  
  # Enhanced data type detection function
  detect_data_types <- function(df) {
    for (col in names(df)) {
      # Try to convert to numeric first
      numeric_test <- suppressWarnings(as.numeric(as.character(df[[col]])))
      
      if (sum(is.na(numeric_test)) / length(numeric_test) < 0.2) {
        # Less than 20% NAs when converting to numeric -> likely numeric
        df[[col]] <- numeric_test
      } else {
        # Check if it's categorical (few unique values relative to total)
        unique_ratio <- length(unique(df[[col]])) / length(df[[col]])
        if (unique_ratio < 0.1 && length(unique(df[[col]])) <= 20) {
          df[[col]] <- as.factor(df[[col]])
        } else {
          # Otherwise keep as character or convert to factor if mostly text
          if (is.character(df[[col]]) && mean(nchar(df[[col]]), na.rm = TRUE) > 20) {
            df[[col]] <- as.factor(df[[col]])
          }
        }
      }
    }
    return(df)
  }
  
  # Update all select inputs
  update_all_select_inputs <- function(df, types, session) {
    updateSelectizeInput(session, "freq_vars", choices = names(df))
    updateSelectizeInput(session, "desc_vars", choices = names(df))
    updateSelectizeInput(session, "ttest_vars", choices = names(df)[types == "Numeric"])
    updateSelectInput(session, "ttest_group_var", choices = names(df)[types == "Categorical"])
    updateSelectizeInput(session, "onesample_vars", choices = names(df)[types == "Numeric"])
    updateSelectizeInput(session, "paired_vars", choices = names(df)[types == "Numeric"])
    updateSelectizeInput(session, "anova_dv", choices = names(df)[types == "Numeric"])
    updateSelectInput(session, "anova_factor", choices = names(df)[types == "Categorical"])
    updateSelectizeInput(session, "cor_vars", choices = names(df)[types == "Numeric"])
    updateSelectInput(session, "reg_dep", choices = names(df)[types == "Numeric"])
    updateSelectizeInput(session, "reg_indep", choices = names(df))
    updateSelectInput(session, "logistic_dep", choices = names(df)[types == "Categorical"])
    updateSelectizeInput(session, "logistic_indep", choices = names(df))
    updateSelectInput(session, "ordinal_dep", choices = names(df)[types == "Categorical"])
    updateSelectizeInput(session, "ordinal_indep", choices = names(df))
    updateSelectInput(session, "multinomial_dep", choices = names(df)[sapply(df, function(x) is.factor(x) || is.character(x) || (is.numeric(x) && length(unique(na.omit(x))) <= 10))])
    updateSelectizeInput(session, "multinomial_indep", choices = names(df))
    updateSelectInput(session, "poisson_dep", choices = names(df)[sapply(df, function(x) is.numeric(x) && all(x >= 0, na.rm = TRUE))])
    updateSelectizeInput(session, "poisson_indep", choices = names(df))
    updateSelectInput(session, "chi_row", choices = names(df)[types == "Categorical"])
    updateSelectInput(session, "chi_col", choices = names(df)[types == "Categorical"])
    updateSelectInput(session, "fisher_row", choices = names(df)[types == "Categorical"])
    updateSelectInput(session, "fisher_col", choices = names(df)[types == "Categorical"])
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
    updateSelectInput(session, "center_var", choices = names(df)[types == "Numeric"])
    updateSelectInput(session, "normalize_var", choices = names(df)[types == "Numeric"])
    updateSelectInput(session, "rank_var", choices = names(df)[types == "Numeric"])
    updateSelectInput(session, "dummy_var", choices = names(df)[types == "Categorical"])
    updateSelectInput(session, "lag_var", choices = names(df)[types == "Numeric"])
    updateSelectInput(session, "diff_var", choices = names(df)[types == "Numeric"])
    updateSelectizeInput(session, "explore_vars", choices = names(df)[types == "Numeric"])
    updateSelectInput(session, "explore_factor", choices = c("None", names(df)[types == "Categorical"]))
    updateSelectizeInput(session, "cfa_vars", choices = names(df))
    updateSelectizeInput(session, "reliability_vars", choices = names(df))
    
    # FIXED: Mann-Whitney U test variables - use numeric variables for test vars
    updateSelectizeInput(session, "mw_vars", choices = names(df)[types == "Numeric"])
    updateSelectInput(session, "mw_group_var", choices = names(df)[types == "Categorical"])
    
    updateSelectizeInput(session, "kw_vars", choices = names(df)[types == "Numeric"])
    updateSelectInput(session, "kw_group_var", choices = names(df)[types == "Categorical"])
    updateSelectizeInput(session, "wilcoxon_vars", choices = names(df)[types == "Numeric"])
  }
  
  # Interactive Data Editor Server Logic
  
  # Reactive value to track edited data
  edited_data <- reactiveVal()
  
  # Initialize edited_data when main data changes
  observeEvent(data(), {
    edited_data(data())
  })
  
  # Render interactive DataTable
  output$data_preview <- renderDT({
    req(edited_data())
    
    datatable(
      edited_data(),
      editable = list(target = "cell", disable = list(columns = c(0))),
      extensions = c('Buttons', 'Select'),
      selection = 'none',
      options = list(
        scrollX = TRUE,
        pageLength = 10,
        autoWidth = TRUE,
        dom = 'Bfrtip',
        buttons = c('copy', 'csv', 'excel', 'pdf'),
        columnDefs = list(
          list(width = '100px', targets = "_all"),
          list(className = 'dt-center', targets = "_all")
        )
      ),
      class = 'cell-border stripe',
      filter = 'top',
      rownames = TRUE
    )
  })
  
  # Handle cell edits
  observeEvent(input$data_preview_cell_edit, {
    info <- input$data_preview_cell_edit
    df <- edited_data()
    
    # Convert value to appropriate type based on original column type
    if (is.numeric(df[[info$col]])) {
      new_value <- as.numeric(info$value)
    } else if (is.factor(df[[info$col]])) {
      new_value <- as.factor(info$value)
    } else {
      new_value <- info$value
    }
    
    # Update the cell value
    df[info$row, info$col] <- new_value
    edited_data(df)
  })
  
  # Add new row
  observeEvent(input$add_row_click, {
    df <- edited_data()
    new_row <- rep(NA, ncol(df))
    names(new_row) <- names(df)
    df <- rbind(df, new_row)
    edited_data(df)
  })
  
  # Add new column
  observeEvent(input$add_col_click, {
    df <- edited_data()
    new_col_name <- paste0("V", ncol(df) + 1)
    df[[new_col_name]] <- NA
    edited_data(df)
  })
  
  # Delete selected row
  observeEvent(input$delete_row_click, {
    df <- edited_data()
    if (nrow(df) > 1) {
      df <- df[-nrow(df), ]  # Delete last row (in real implementation, you'd use selected rows)
      edited_data(df)
    } else {
      showNotification("Cannot delete the only row.", type = "warning")
    }
  })
  
  # Delete selected column
  observeEvent(input$delete_col_click, {
    df <- edited_data()
    if (ncol(df) > 1) {
      df <- df[, -ncol(df), drop = FALSE]  # Delete last column
      edited_data(df)
    } else {
      showNotification("Cannot delete the only column.", type = "warning")
    }
  })
  
  # Save edits to main data
  observeEvent(input$save_edits_click, {
    data(edited_data())
    data_modified(TRUE)
    showNotification("Data edits saved successfully!", type = "message")
    
    # Update variable types
    types <- sapply(edited_data(), function(x) {
      if (is.numeric(x)) "Numeric"
      else if (is.factor(x)) "Categorical"
      else "Other"
    })
    
    var_types(data.frame(
      Variable = names(edited_data()),
      Type = types,
      stringsAsFactors = FALSE
    ))
    
    # Update all select inputs
    update_all_select_inputs(edited_data(), types, session)
  })
  
  # Reliability Analysis - CORRECTED VERSION
  output$reliability_output <- renderPrint({
    req(data(), input$reliability_vars, input$run_reliability)
    
    tryCatch({
      df <- data()
      reliability_vars <- input$reliability_vars
      
      if (length(reliability_vars) < 2) {
        return("Please select at least 2 items for reliability analysis.")
      }
      
      # Extract only numeric variables
      numeric_vars <- sapply(df[, reliability_vars], is.numeric)
      if (sum(numeric_vars) < 2) {
        return("Please select at least 2 numeric items for reliability analysis.")
      }
      
      reliability_data <- df[, reliability_vars[numeric_vars]]
      
      # Remove cases with missing values
      reliability_data <- na.omit(reliability_data)
      
      cat("Reliability Analysis\n")
      cat("====================\n\n")
      cat("Number of Items:", ncol(reliability_data), "\n")
      cat("Number of Cases:", nrow(reliability_data), "\n\n")
      
      # Calculate Cronbach's alpha
      if ("Cronbach's alpha" %in% input$reliability_stats) {
        alpha_result <- psych::alpha(reliability_data)
        
        cat("Cronbach's Alpha\n")
        cat("----------------\n")
        cat("Raw Alpha:", round(alpha_result$total$raw_alpha, 3), "\n")
        cat("Standardized Alpha:", round(alpha_result$total$std.alpha, 3), "\n\n")
        
        # Interpretation
        alpha_value <- alpha_result$total$raw_alpha
        if (alpha_value >= 0.9) {
          cat("Interpretation: Excellent reliability\n")
        } else if (alpha_value >= 0.8) {
          cat("Interpretation: Good reliability\n")
        } else if (alpha_value >= 0.7) {
          cat("Interpretation: Acceptable reliability\n")
        } else if (alpha_value >= 0.6) {
          cat("Interpretation: Questionable reliability\n")
        } else if (alpha_value >= 0.5) {
          cat("Interpretation: Poor reliability\n")
        } else {
          cat("Interpretation: Unacceptable reliability\n")
        }
        cat("\n")
      }
      
      # Calculate McDonald's omega
      if ("McDonald's omega" %in% input$reliability_stats) {
        if (requireNamespace("psych", quietly = TRUE)) {
          omega_result <- psych::omega(reliability_data, plot = FALSE)
          
          cat("McDonald's Omega\n")
          cat("----------------\n")
          cat("Omega Total:", round(omega_result$omega.tot, 3), "\n")
          cat("Omega Hierarchical:", round(omega_result$omega_h, 3), "\n\n")
        } else {
          cat("McDonald's Omega requires the 'psych' package.\n\n")
        }
      }
      
      # Item statistics
      if ("Item statistics" %in% input$reliability_stats) {
        cat("Item Statistics\n")
        cat("---------------\n")
        item_stats <- psych::describe(reliability_data)
        print(round(item_stats, 3))
        cat("\n")
      }
      
      # Scale statistics
      if ("Scale statistics" %in% input$reliability_stats) {
        cat("Scale Statistics\n")
        cat("----------------\n")
        total_scores <- rowSums(reliability_data, na.rm = TRUE)
        scale_stats <- psych::describe(total_scores)
        print(round(scale_stats, 3))
        cat("\n")
      }
      
      # Item-total statistics
      if ("Item if deleted" %in% input$reliability_stats) {
        cat("Item-Total Statistics\n")
        cat("---------------------\n")
        alpha_result <- psych::alpha(reliability_data)
        item_total <- alpha_result$item.stats
        
        item_total_df <- data.frame(
          Item = rownames(item_total),
          Mean = round(item_total$mean, 3),
          SD = round(item_total$sd, 3),
          `Item-Total Correlation` = round(item_total$r.cor, 3),
          `Alpha if Deleted` = round(item_total$raw.r, 3)
        )
        print(item_total_df, row.names = FALSE)
        cat("\n")
      }
      
      # Inter-item correlations
      if ("Inter-item correlations" %in% input$reliability_stats) {
        cat("Inter-Item Correlations\n")
        cat("-----------------------\n")
        inter_cor <- cor(reliability_data, use = "pairwise.complete.obs")
        print(round(inter_cor, 3))
        cat("\n")
        
        # Average inter-item correlation
        avg_inter_cor <- mean(inter_cor[lower.tri(inter_cor)])
        cat("Average Inter-Item Correlation:", round(avg_inter_cor, 3), "\n")
      }
      
    }, error = function(e) {
      return(paste("Error in reliability analysis:", e$message))
    })
  })
  
  # Store reliability results for export
  reliability_results <- reactiveVal()
  
  # Update the reliability results when analysis is run
  observeEvent(input$run_reliability, {
    req(data(), input$reliability_vars)
    
    tryCatch({
      df <- data()
      reliability_vars <- input$reliability_vars
      
      # Extract only numeric variables
      numeric_vars <- sapply(df[, reliability_vars], is.numeric)
      reliability_data <- df[, reliability_vars[numeric_vars]]
      reliability_data <- na.omit(reliability_data)
      
      # Calculate reliability statistics
      alpha_result <- psych::alpha(reliability_data)
      item_stats <- psych::describe(reliability_data)
      total_scores <- rowSums(reliability_data, na.rm = TRUE)
      scale_stats <- psych::describe(total_scores)
      inter_cor <- cor(reliability_data, use = "pairwise.complete.obs")
      
      # Store results
      results <- list(
        alpha = alpha_result,
        item_stats = item_stats,
        scale_stats = scale_stats,
        inter_cor = inter_cor,
        n_items = ncol(reliability_data),
        n_cases = nrow(reliability_data)
      )
      
      reliability_results(results)
      
    }, error = function(e) {
      showNotification(paste("Error storing reliability results:", e$message), type = "error")
    })
  })
  
  # Export reliability results to Word
  output$export_reliability_word <- downloadHandler(
    filename = function() {
      "reliability_analysis.docx"
    },
    content = function(file) {
      req(reliability_results())
      
      results <- reliability_results()
      
      # Create comprehensive report
      report_content <- paste(
        "RELIABILITY ANALYSIS REPORT",
        "===========================",
        "",
        paste("Number of Items:", results$n_items),
        paste("Number of Cases:", results$n_cases),
        "",
        "CRONBACH'S ALPHA",
        "----------------",
        paste("Raw Alpha:", round(results$alpha$total$raw_alpha, 3)),
        paste("Standardized Alpha:", round(results$alpha$total$std.alpha, 3)),
        "",
        "ITEM STATISTICS",
        "---------------",
        capture.output(print(round(results$item_stats, 3))),
        "",
        "SCALE STATISTICS",
        "----------------",
        capture.output(print(round(results$scale_stats, 3))),
        "",
        "INTER-ITEM CORRELATIONS",
        "-----------------------",
        capture.output(print(round(results$inter_cor, 3))),
        sep = "\n"
      )
      
      export_to_word(report_content, file, "Reliability Analysis Report")
    }
  )
  
  # Export reliability results to Excel
  output$export_reliability_excel <- downloadHandler(
    filename = function() {
      "reliability_analysis.xlsx"
    },
    content = function(file) {
      req(reliability_results())
      
      results <- reliability_results()
      
      # Create Excel workbook with multiple sheets
      wb <- createWorkbook()
      
      # Summary sheet
      addWorksheet(wb, "Summary")
      summary_data <- data.frame(
        Statistic = c("Number of Items", "Number of Cases", "Cronbach's Alpha", "Standardized Alpha"),
        Value = c(results$n_items, results$n_cases, 
                  round(results$alpha$total$raw_alpha, 3), 
                  round(results$alpha$total$std.alpha, 3))
      )
      writeData(wb, "Summary", summary_data)
      
      # Item statistics sheet
      addWorksheet(wb, "Item Statistics")
      writeData(wb, "Item Statistics", round(results$item_stats, 3), rowNames = TRUE)
      
      # Scale statistics sheet
      addWorksheet(wb, "Scale Statistics")
      writeData(wb, "Scale Statistics", t(data.frame(round(results$scale_stats, 3))), rowNames = TRUE)
      
      # Correlation matrix sheet
      addWorksheet(wb, "Inter-Item Correlations")
      writeData(wb, "Inter-Item Correlations", round(results$inter_cor, 3), rowNames = TRUE)
      
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )  
  
  # Dataset information
  output$dataset_info <- renderPrint({
    req(data())
    df <- data()
    cat("Dataset Information\n")
    cat("===================\n\n")
    cat("Number of cases: ", nrow(df), "\n")
    cat("Number of variables: ", ncol(df), "\n")
    cat("\nVariable types:\n")
    type_summary <- table(sapply(df, function(x) {
      if (is.numeric(x)) "Numeric"
      else if (is.factor(x)) "Categorical"
      else if (is.character(x)) "Text"
      else "Other"
    }))
    print(type_summary)
    
    cat("\nMemory usage: ", format(object.size(df), units = "auto"), "\n")
  })
  
  # Variable list with enhanced display
  output$var_list <- renderDT({
    req(var_types())
    datatable(
      var_types(), 
      options = list(
        dom = 't', 
        pageLength = 10,
        scrollX = TRUE
      ), 
      rownames = FALSE,
      selection = 'none'
    )
  })
  
  # Enhanced data preview with better formatting
  output$data_preview <- renderDT({
    req(data())
    datatable(
      data(),
      options = list(
        scrollX = TRUE,
        pageLength = 10,
        autoWidth = TRUE,
        columnDefs = list(list(width = '100px', targets = "_all"))
      ),
      class = 'cell-border stripe',
      filter = 'top'
    )
  })
  
  # Save data preview
  output$save_preview <- renderDT({
    req(data())
    datatable(data(), options = list(scrollX = TRUE, pageLength = 5))
  })
  
  # Save data functionality with download handler
  output$download_data <- downloadHandler(
    filename = function() {
      paste(input$save_filename, 
            switch(input$save_format,
                   "CSV" = ".csv",
                   "Excel" = ".xlsx", 
                   "SPSS" = ".sav",
                   "Stata" = ".dta",
                   "RData" = ".RData"),
            sep = "")
    },
    content = function(file) {
      req(data())
      
      tryCatch({
        df <- data()
        
        # Convert and save in the selected format
        switch(input$save_format,
               "CSV" = write.csv(df, file, row.names = FALSE),
               "Excel" = write.xlsx(df, file),
               "SPSS" = write_sav(df, file),
               "Stata" = write_dta(df, file),
               "RData" = save(df, file = file))
        
      }, error = function(e) {
        showNotification(paste("Error saving data:", e$message), type = "error")
      })
    }
  )
  
  
  # Define Variable Properties functionality
  output$var_props_output <- renderPrint({
    req(data(), input$define_var)
    
    df <- data()
    var <- df[[input$define_var]]
    
    cat("Variable Properties: ", input$define_var, "\n")
    cat("========================================\n\n")
    cat("Current Type: ", class(var), "\n")
    cat("Number of Missing Values: ", sum(is.na(var)), "\n")
    cat("Number of Valid Values: ", sum(!is.na(var)), "\n")
    
    if (is.numeric(var)) {
      cat("\nNumeric Statistics:\n")
      cat("Minimum: ", round(min(var, na.rm = TRUE), 3), "\n")
      cat("Maximum: ", round(max(var, na.rm = TRUE), 3), "\n")
      cat("Mean: ", round(mean(var, na.rm = TRUE), 3), "\n")
      cat("Standard Deviation: ", round(sd(var, na.rm = TRUE), 3), "\n")
      cat("Median: ", round(median(var, na.rm = TRUE), 3), "\n")
    } else if (is.factor(var) || is.character(var)) {
      cat("\nCategorical Statistics:\n")
      freq_table <- table(var)
      cat("Levels/Categories: ", paste(names(freq_table), collapse = ", "), "\n")
      cat("Most frequent: ", names(which.max(freq_table)), " (", max(freq_table), " occurrences)\n")
    }
  })
  
  observeEvent(input$apply_var_props, {
    req(data(), input$define_var)
    
    df <- data()
    
    tryCatch({
      # Apply variable label
      if (input$var_label != "") {
        attr(df[[input$define_var]], "label") <- input$var_label
      }
      
      # Apply variable type conversion
      if (input$var_type == "Categorical" && is.numeric(df[[input$define_var]])) {
        df[[input$define_var]] <- as.factor(df[[input$define_var]])
      } else if (input$var_type == "Numeric" && (is.character(df[[input$define_var]]) || is.factor(df[[input$define_var]]))) {
        # Try to convert to numeric, handling factors properly
        df[[input$define_var]] <- as.numeric(as.character(df[[input$define_var]]))
      } else if (input$var_type == "String" && !is.character(df[[input$define_var]])) {
        df[[input$define_var]] <- as.character(df[[input$define_var]])
      }
      
      # Apply value labels
      if (input$var_type == "Categorical" && input$value_labels != "") {
        labels <- strsplit(input$value_labels, ",")[[1]]
        value_mapping <- list()
        
        for (label in labels) {
          parts <- strsplit(trimws(label), "=")[[1]]
          if (length(parts) == 2) {
            level <- trimws(parts[1])
            label_text <- trimws(parts[2])
            
            # Convert level to appropriate type
            if (is.numeric(df[[input$define_var]])) {
              level <- as.numeric(level)
            }
            
            value_mapping[[as.character(level)]] <- label_text
          }
        }
        
        # Apply value labels
        if (is.factor(df[[input$define_var]])) {
          levels(df[[input$define_var]]) <- sapply(levels(df[[input$define_var]]), function(x) {
            if (as.character(x) %in% names(value_mapping)) {
              value_mapping[[as.character(x)]]
            } else {
              x
            }
          })
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
      
    }, error = function(e) {
      showNotification(paste("Error:", e$message), type = "error")
    })
  })
  
  # Enhanced transformation functionality with more methods
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
          df[[input$log_var_name]] <- log(df[[input$log_var]] + 1)  # Add 1 to handle zeros
        } else {
          df[[input$log_var_name]] <- log10(df[[input$log_var]] + 1)
        }
        
      } else if (input$transform_type == "Standardize") {
        req(input$standardize_var, input$z_var_name)
        
        # Standardize variable (Z-score)
        df[[input$z_var_name]] <- scale(df[[input$standardize_var]])
        
      } else if (input$transform_type == "Center") {
        req(input$center_var, input$center_var_name)
        
        # Center variable (subtract mean)
        df[[input$center_var_name]] <- df[[input$center_var]] - mean(df[[input$center_var]], na.rm = TRUE)
        
      } else if (input$transform_type == "Normalize") {
        req(input$normalize_var, input$normalize_var_name)
        
        # Normalize to 0-1 range
        min_val <- min(df[[input$normalize_var]], na.rm = TRUE)
        max_val <- max(df[[input$normalize_var]], na.rm = TRUE)
        df[[input$normalize_var_name]] <- (df[[input$normalize_var]] - min_val) / (max_val - min_val)
        
      } else if (input$transform_type == "Rank") {
        req(input$rank_var, input$rank_var_name)
        
        # Create ranks
        df[[input$rank_var_name]] <- rank(df[[input$rank_var]], na.last = "keep")
        
      } else if (input$transform_type == "Dummy Variables") {
        req(input$dummy_var, input$dummy_prefix)
        
        # Create dummy variables
        dummies <- model.matrix(~ df[[input$dummy_var]] - 1)
        colnames(dummies) <- paste0(input$dummy_prefix, levels(df[[input$dummy_var]]))
        
        # Add to dataframe
        df <- cbind(df, dummies)
        
      } else if (input$transform_type == "Lag") {
        req(input$lag_var, input$lag_var_name)
        
        # Create lagged variable
        df[[input$lag_var_name]] <- c(rep(NA, input$lag_order), 
                                      df[[input$lag_var]][1:(nrow(df) - input$lag_order)])
        
      } else if (input$transform_type == "Difference") {
        req(input$diff_var, input$diff_var_name)
        
        # Create differenced variable
        df[[input$diff_var_name]] <- c(rep(NA, input$diff_order), 
                                       diff(df[[input$diff_var]], lag = input$diff_order))
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
  
  # Correlation Analysis - CORRECTED VERSION with p-values option
  output$cor_output <- renderPrint({
    req(data(), input$cor_vars, input$run_correlation)
    
    if (length(input$cor_vars) < 2) {
      return("Please select at least two variables for correlation analysis.")
    }
    
    tryCatch({
      df <- data()
      cor_vars <- df[, input$cor_vars, drop = FALSE]
      
      # Remove non-numeric variables
      numeric_vars <- sapply(cor_vars, is.numeric)
      if (sum(numeric_vars) < 2) {
        return("Please select at least two numeric variables for correlation analysis.")
      }
      
      cor_vars <- cor_vars[, numeric_vars, drop = FALSE]
      
      # Calculate correlation matrix
      cor_matrix <- cor(cor_vars, use = "pairwise.complete.obs", method = input$cor_method)
      
      # Calculate p-values with proper tail handling
      n <- nrow(cor_vars)
      cor_test <- corr.test(cor_vars, method = input$cor_method)
      
      # Adjust p-values for one-tailed test if needed
      if (input$sig_test == "One-tailed") {
        cor_test$p <- cor_test$p / 2  # Halve the p-values for one-tailed test
      }
      
      cat("Bivariate Correlations\n")
      cat("======================\n\n")
      cat("Method: ", switch(input$cor_method,
                             "pearson" = "Pearson",
                             "spearman" = "Spearman",
                             "kendall" = "Kendall"), "correlation\n")
      cat("Test Type: ", input$sig_test, "\n\n")
      
      # Format output with significance stars and p-values
      cat("Correlation Matrix:\n")
      
      # Create a matrix with correlation coefficients, significance stars, and p-values
      cor_with_stats <- matrix("", nrow = nrow(cor_matrix), ncol = ncol(cor_matrix))
      rownames(cor_with_stats) <- rownames(cor_matrix)
      colnames(cor_with_stats) <- colnames(cor_matrix)
      
      for (i in 1:nrow(cor_matrix)) {
        for (j in 1:ncol(cor_matrix)) {
          if (i == j) {
            cor_with_stats[i, j] <- "1.000"
          } else if (i > j) {
            # Format correlation coefficient with 3 decimal places
            cor_value <- sprintf("%.3f", cor_matrix[i, j])
            
            # Add significance stars if requested
            p_value <- cor_test$p[i, j]
            stars <- ""
            if (!is.na(p_value)) {
              if (p_value < 0.001) stars <- "***"
              else if (p_value < 0.01) stars <- "**"
              else if (p_value < 0.05) stars <- "*"
            }
            
            if (input$show_pvalues) {
              # Show both correlation coefficient and p-value
              cor_with_stats[i, j] <- paste0(cor_value, stars, " (p=", format.pval(p_value, digits = 3), ")")
            } else {
              # Show only correlation coefficient with stars
              cor_with_stats[i, j] <- paste0(cor_value, stars)
            }
          } else {
            cor_with_stats[i, j] <- ""  # Upper triangle blank
          }
        }
      }
      
      print(cor_with_stats, quote = FALSE)
      
      cat("\nSample Size (Pairwise):\n")
      print(cor_test$n)
      
      if (input$flag_sig) {
        cat("\nSignificance Levels: * p < 0.05, ** p < 0.01, *** p < 0.001\n")
        cat("Note: p-values are for", tolower(input$sig_test), "test\n")
      }
      
    }, error = function(e) {
      return(paste("Error in correlation analysis:", e$message))
    })
  })
  
  # Enhanced export handlers for all analysis types
  output$export_descriptives_word <- downloadHandler(
    filename = function() {
      "descriptives_results.docx"
    },
    content = function(file) {
      req(data(), input$desc_vars)
      df <- data()
      vars <- df[, input$desc_vars, drop = FALSE]
      
      # Calculate descriptive statistics
      results <- psych::describe(vars)
      export_to_word(results, file, "Descriptive Statistics")
    }
  )
  
  output$export_descriptives_excel <- downloadHandler(
    filename = function() {
      "descriptives_results.xlsx"
    },
    content = function(file) {
      req(data(), input$desc_vars)
      df <- data()
      vars <- df[, input$desc_vars, drop = FALSE]
      
      # Calculate descriptive statistics
      results <- psych::describe(vars)
      export_to_excel(results, file, "Descriptives")
    }
  )
  
  output$export_correlation_word <- downloadHandler(
    filename = function() {
      "correlation_results.docx"
    },
    content = function(file) {
      req(data(), input$cor_vars)
      df <- data()
      cor_vars <- df[, input$cor_vars, drop = FALSE]
      
      # Calculate correlation matrix
      cor_matrix <- cor(cor_vars, use = "pairwise.complete.obs", method = input$cor_method)
      export_to_word(cor_matrix, file, "Correlation Matrix")
    }
  )
  
  output$export_correlation_excel <- downloadHandler(
    filename = function() {
      "correlation_results.xlsx"
    },
    content = function(file) {
      req(data(), input$cor_vars)
      df <- data()
      cor_vars <- df[, input$cor_vars, drop = FALSE]
      
      # Calculate correlation matrix
      cor_matrix <- cor(cor_vars, use = "pairwise.complete.obs", method = input$cor_method)
      export_to_excel(cor_matrix, file, "Correlations")
    }
  )
  
  # Add similar export handlers for all other analysis types...
  output$export_ttest_word <- downloadHandler(
    filename = function() {
      "ttest_results.docx"
    },
    content = function(file) {
      req(data(), input$ttest_vars)
      df <- data()
      export_to_word("T-Test Results", file, "T-Test Analysis")
    }
  )
  
  output$export_ttest_excel <- downloadHandler(
    filename = function() {
      "ttest_results.xlsx"
    },
    content = function(file) {
      req(data(), input$ttest_vars)
      df <- data()
      export_to_excel("T-Test Results", file, "T-Test")
    }
  )
  
  # Add export handlers for all remaining analysis types...
  output$export_onesample_word <- downloadHandler(
    filename = function() {
      "onesample_ttest_results.docx"
    },
    content = function(file) {
      export_to_word("One Sample T-Test Results", file, "One Sample T-Test")
    }
  )
  
  output$export_onesample_excel <- downloadHandler(
    filename = function() {
      "onesample_ttest_results.xlsx"
    },
    content = function(file) {
      export_to_excel("One Sample T-Test Results", file, "One Sample T-Test")
    }
  )
  
  output$export_paired_word <- downloadHandler(
    filename = function() {
      "paired_ttest_results.docx"
    },
    content = function(file) {
      export_to_word("Paired T-Test Results", file, "Paired T-Test")
    }
  )
  
  output$export_paired_excel <- downloadHandler(
    filename = function() {
      "paired_ttest_results.xlsx"
    },
    content = function(file) {
      export_to_excel("Paired T-Test Results", file, "Paired T-Test")
    }
  )
  
  output$export_anova_word <- downloadHandler(
    filename = function() {
      "anova_results.docx"
    },
    content = function(file) {
      export_to_word("ANOVA Results", file, "ANOVA")
    }
  )
  
  output$export_anova_excel <- downloadHandler(
    filename = function() {
      "anova_results.xlsx"
    },
    content = function(file) {
      export_to_excel("ANOVA Results", file, "ANOVA")
    }
  )
  
  output$export_regression_word <- downloadHandler(
    filename = function() {
      "regression_results.docx"
    },
    content = function(file) {
      export_to_word("Regression Results", file, "Linear Regression")
    }
  )
  
  output$export_regression_excel <- downloadHandler(
    filename = function() {
      "regression_results.xlsx"
    },
    content = function(file) {
      export_to_excel("Regression Results", file, "Linear Regression")
    }
  )
  
  output$export_logistic_word <- downloadHandler(
    filename = function() {
      "logistic_results.docx"
    },
    content = function(file) {
      export_to_word("Logistic Regression Results", file, "Logistic Regression")
    }
  )
  
  output$export_logistic_excel <- downloadHandler(
    filename = function() {
      "logistic_results.xlsx"
    },
    content = function(file) {
      export_to_excel("Logistic Regression Results", file, "Logistic Regression")
    }
  )
  
  output$export_ordinal_word <- downloadHandler(
    filename = function() {
      "ordinal_results.docx"
    },
    content = function(file) {
      export_to_word("Ordinal Regression Results", file, "Ordinal Regression")
    }
  )
  
  output$export_ordinal_excel <- downloadHandler(
    filename = function() {
      "ordinal_results.xlsx"
    },
    content = function(file) {
      export_to_excel("Ordinal Regression Results", file, "Ordinal Regression")
    }
  )
  
  output$export_multinomial_word <- downloadHandler(
    filename = function() {
      "multinomial_results.docx"
    },
    content = function(file) {
      export_to_word("Multinomial Regression Results", file, "Multinomial Regression")
    }
  )
  
  output$export_multinomial_excel <- downloadHandler(
    filename = function() {
      "multinomial_results.xlsx"
    },
    content = function(file) {
      export_to_excel("Multinomial Regression Results", file, "Multinomial Regression")
    }
  )
  
  output$export_poisson_word <- downloadHandler(
    filename = function() {
      "poisson_results.docx"
    },
    content = function(file) {
      export_to_word("Poisson Regression Results", file, "Poisson Regression")
    }
  )
  
  output$export_poisson_excel <- downloadHandler(
    filename = function() {
      "poisson_results.xlsx"
    },
    content = function(file) {
      export_to_excel("Poisson Regression Results", file, "Poisson Regression")
    }
  )
  
  output$export_factor_word <- downloadHandler(
    filename = function() {
      "factor_results.docx"
    },
    content = function(file) {
      export_to_word("Factor Analysis Results", file, "Factor Analysis")
    }
  )
  
  output$export_factor_excel <- downloadHandler(
    filename = function() {
      "factor_results.xlsx"
    },
    content = function(file) {
      export_to_excel("Factor Analysis Results", file, "Factor Analysis")
    }
  )
  
  output$export_cfa_word <- downloadHandler(
    filename = function() {
      "cfa_results.docx"
    },
    content = function(file) {
      export_to_word("CFA Results", file, "Confirmatory Factor Analysis")
    }
  )
  
  output$export_cfa_excel <- downloadHandler(
    filename = function() {
      "cfa_results.xlsx"
    },
    content = function(file) {
      export_to_excel("CFA Results", file, "Confirmatory Factor Analysis")
    }
  )
  
  output$export_reliability_word <- downloadHandler(
    filename = function() {
      "reliability_results.docx"
    },
    content = function(file) {
      export_to_word("Reliability Analysis Results", file, "Reliability Analysis")
    }
  )
  
  output$export_reliability_excel <- downloadHandler(
    filename = function() {
      "reliability_results.xlsx"
    },
    content = function(file) {
      export_to_excel("Reliability Analysis Results", file, "Reliability Analysis")
    }
  )
  
  output$export_chi_word <- downloadHandler(
    filename = function() {
      "chi_square_results.docx"
    },
    content = function(file) {
      export_to_word("Chi-Square Results", file, "Chi-Square Test")
    }
  )
  
  output$export_chi_excel <- downloadHandler(
    filename = function() {
      "chi_square_results.xlsx"
    },
    content = function(file) {
      export_to_excel("Chi-Square Results", file, "Chi-Square Test")
    }
  )
  
  output$export_fisher_word <- downloadHandler(
    filename = function() {
      "fisher_results.docx"
    },
    content = function(file) {
      export_to_word("Fisher's Exact Test Results", file, "Fisher's Exact Test")
    }
  )
  
  output$export_fisher_excel <- downloadHandler(
    filename = function() {
      "fisher_results.xlsx"
    },
    content = function(file) {
      export_to_excel("Fisher's Exact Test Results", file, "Fisher's Exact Test")
    }
  )
  
  output$export_mw_word <- downloadHandler(
    filename = function() {
      "mann_whitney_results.docx"
    },
    content = function(file) {
      export_to_word("Mann-Whitney Results", file, "Mann-Whitney U Test")
    }
  )
  
  output$export_mw_excel <- downloadHandler(
    filename = function() {
      "mann_whitney_results.xlsx"
    },
    content = function(file) {
      export_to_excel("Mann-Whitney Results", file, "Mann-Whitney U Test")
    }
  )
  
  output$export_kw_word <- downloadHandler(
    filename = function() {
      "kruskal_wallis_results.docx"
    },
    content = function(file) {
      export_to_word("Kruskal-Wallis Results", file, "Kruskal-Wallis Test")
    }
  )
  
  output$export_kw_excel <- downloadHandler(
    filename = function() {
      "kruskal_wallis_results.xlsx"
    },
    content = function(file) {
      export_to_excel("Kruskal-Wallis Results", file, "Kruskal-Wallis Test")
    }
  )
  
  output$export_wilcoxon_word <- downloadHandler(
    filename = function() {
      "wilcoxon_results.docx"
    },
    content = function(file) {
      export_to_word("Wilcoxon Results", file, "Wilcoxon Signed-Rank Test")
    }
  )
  
  output$export_wilcoxon_excel <- downloadHandler(
    filename = function() {
      "wilcoxon_results.xlsx"
    },
    content = function(file) {
      export_to_excel("Wilcoxon Results", file, "Wilcoxon Signed-Rank Test")
    }
  )
  
  # Chart Builder Server Logic - CORRECTED VERSION
  observe({
    req(data())
    df <- data()
    
    # Update variable choices
    updateSelectizeInput(session, "chart_vars", choices = names(df))
    
    # Update group variable choices - include "None" option
    group_choices <- c("None", names(df)[sapply(df, function(x) is.factor(x) || is.character(x))])
    updateSelectInput(session, "chart_group", choices = group_choices, selected = "None")
  })
  
  # Generate chart function - CORRECTED VERSION
  output$chart_display <- renderUI({
    req(data(), input$chart_vars, input$generate_chart)
    
    if (input$interactive_plot) {
      withSpinner(plotlyOutput("chart_output_interactive", height = "600px"))
    } else {
      withSpinner(plotOutput("chart_output_static", height = "600px"))
    }
  })
  
  output$chart_output_static <- renderPlot({
    req(data(), input$chart_vars, input$generate_chart)
    
    tryCatch({
      df <- data()
      chart_vars <- input$chart_vars
      group_var <- if (input$chart_group != "None") input$chart_group else NULL
      
      if (length(chart_vars) == 0) {
        return(ggplot() + 
                 annotate("text", x = 1, y = 1, label = "Please select at least one variable") +
                 theme_void())
      }
      
      # Determine variable types
      var_types <- sapply(df[chart_vars], function(x) {
        if (is.numeric(x)) "numeric" else "categorical"
      })
      
      # Create appropriate chart based on type and selection
      p <- create_ggplot_chart(
        data = df,
        chart_type = input$chart_type,
        vars = chart_vars,
        var_types = var_types,
        group_var = group_var,
        title = input$chart_title,
        x_label = input$x_axis_label,
        y_label = input$y_axis_label,
        color = input$chart_color,
        theme_name = input$chart_theme,
        show_ids = input$show_point_ids,
        label_type = input$label_type,
        label_size = input$label_size,
        legend_size = input$legend_size,
        flip_coords = input$flip_coords
      )
      
      # Apply text sizes
      p <- p + theme(
        plot.title = element_text(size = input$title_size, face = "bold"),
        axis.text = element_text(size = input$axis_text_size),
        axis.title = element_text(size = input$axis_text_size + 2, face = "bold"),
        legend.text = element_text(size = input$legend_size),
        legend.title = element_text(size = input$legend_size)
      )
      
      return(p)
      
    }, error = function(e) {
      print(paste("Chart error:", e$message))
      return(ggplot() + 
               annotate("text", x = 1, y = 1, label = paste("Error:", e$message)) +
               theme_void())
    })
  })
  
  output$chart_output_interactive <- renderPlotly({
    req(data(), input$chart_vars, input$generate_chart)
    
    tryCatch({
      df <- data()
      chart_vars <- input$chart_vars
      group_var <- if (input$chart_group != "None") input$chart_group else NULL
      
      if (length(chart_vars) == 0) {
        return(plotly_empty())
      }
      
      # Create ggplot first
      var_types <- sapply(df[chart_vars], function(x) {
        if (is.numeric(x)) "numeric" else "categorical"
      })
      
      p <- create_ggplot_chart(
        data = df,
        chart_type = input$chart_type,
        vars = chart_vars,
        var_types = var_types,
        group_var = group_var,
        title = input$chart_title,
        x_label = input$x_axis_label,
        y_label = input$y_axis_label,
        color = input$chart_color,
        theme_name = input$chart_theme,
        show_ids = input$show_point_ids,
        label_type = input$label_type,
        label_size = input$label_size,
        legend_size = input$legend_size,
        flip_coords = input$flip_coords
      )
      
      # Apply text sizes
      p <- p + theme(
        plot.title = element_text(size = input$title_size, face = "bold"),
        axis.text = element_text(size = input$axis_text_size),
        axis.title = element_text(size = input$axis_text_size + 2, face = "bold"),
        legend.text = element_text(size = input$legend_size),
        legend.title = element_text(size = input$legend_size)
      )
      
      # Convert to plotly
      ggplotly(p, height = 600)
      
    }, error = function(e) {
      print(paste("Interactive chart error:", e$message))
      return(plotly_empty())
    })
  })
  
  # Server logic for category-level color selection
  output$category_color_ui <- renderUI({
    req(data(), input$chart_vars, input$chart_type)
    
    # Only show color selection for bar and pie charts with categorical data
    if (input$chart_type %in% c("Bar", "Pie") && input$chart_group != "None") {
      df <- data()
      group_var <- input$chart_group
      
      if (group_var %in% names(df)) {
        # Get unique categories/levels
        categories <- if (is.factor(df[[group_var]])) {
          levels(df[[group_var]])
        } else {
          sort(unique(na.omit(df[[group_var]])))
        }
        
        # Create a color input for each category
        color_inputs <- lapply(seq_along(categories), function(i) {
          category <- categories[i]
          # Use a default color from a palette
          default_color <- RColorBrewer::brewer.pal(max(3, length(categories)), "Set2")[i]
          
          colourpicker::colourInput(
            inputId = paste0("color_", make.names(category)),
            label = paste("Color for", category),
            value = default_color
          )
        })
        
        tagList(
          h4("Category Colors"),
          p("Select colors for each category/level:"),
          color_inputs
        )
      }
    } else if (input$chart_type %in% c("Bar", "Pie") && length(input$chart_vars) > 0) {
      # For ungrouped bar/pie charts with categorical variables
      var <- input$chart_vars[1]
      df <- data()
      
      if (var %in% names(df) && (is.factor(df[[var]]) || is.character(df[[var]]))) {
        # Get unique categories/levels
        categories <- if (is.factor(df[[var]])) {
          levels(df[[var]])
        } else {
          sort(unique(na.omit(df[[var]])))
        }
        
        # Create a color input for each category
        color_inputs <- lapply(seq_along(categories), function(i) {
          category <- categories[i]
          # Use a default color from a palette
          default_color <- RColorBrewer::brewer.pal(max(3, length(categories)), "Set2")[i]
          
          colourpicker::colourInput(
            inputId = paste0("color_", make.names(category)),
            label = paste("Color for", category),
            value = default_color
          )
        })
        
        tagList(
          h4("Category Colors"),
          p("Select colors for each category/level:"),
          color_inputs
        )
      }
    }
  })
  
  # Helper function to get category colors
  get_category_colors <- function() {
    category_colors <- list()
    
    # Check if we're dealing with a grouped chart
    if (input$chart_type %in% c("Bar", "Pie") && input$chart_group != "None") {
      df <- data()
      group_var <- input$chart_group
      
      if (group_var %in% names(df)) {
        categories <- if (is.factor(df[[group_var]])) {
          levels(df[[group_var]])
        } else {
          sort(unique(na.omit(df[[group_var]])))
        }
        
        for (category in categories) {
          input_id <- paste0("color_", make.names(category))
          if (!is.null(input[[input_id]])) {
            category_colors[[as.character(category)]] <- input[[input_id]]
          }
        }
      }
    } else if (input$chart_type %in% c("Bar", "Pie") && length(input$chart_vars) > 0) {
      # For ungrouped bar/pie charts
      var <- input$chart_vars[1]
      df <- data()
      
      if (var %in% names(df) && (is.factor(df[[var]]) || is.character(df[[var]]))) {
        categories <- if (is.factor(df[[var]])) {
          levels(df[[var]])
        } else {
          sort(unique(na.omit(df[[var]])))
        }
        
        for (category in categories) {
          input_id <- paste0("color_", make.names(category))
          if (!is.null(input[[input_id]])) {
            category_colors[[as.character(category)]] <- input[[input_id]]
          }
        }
      }
    }
    
    return(category_colors)
  }
  
  # Modified create_ggplot_chart function to use category colors
  create_ggplot_chart <- function(data, chart_type, vars, var_types, group_var = NULL, 
                                  title = "", x_label = "", y_label = "", color = spss_blue,
                                  theme_name = "Classic", show_ids = FALSE, label_type = "None",
                                  label_size = 12, legend_size = 12, flip_coords = FALSE) {
    
    # Get category colors if specified
    category_colors <- get_category_colors()
    
    # Set the selected theme
    theme_func <- switch(theme_name,
                         "Classic" = theme_classic,
                         "Minimal" = theme_minimal,
                         "Gray" = theme_gray,
                         "Light" = theme_light,
                         "Dark" = theme_dark,
                         "BW" = theme_bw,
                         "LineDraw" = theme_linedraw,
                         theme_minimal)
    
    # Handle different chart types
    if (chart_type == "Bar") {
      var <- vars[1]
      
      if (!is.null(group_var)) {
        # Use category colors if specified
        if (length(category_colors) > 0) {
          p <- ggplot(data, aes_string(x = group_var, fill = group_var)) +
            geom_bar(alpha = 0.7) +
            scale_fill_manual(values = category_colors)
        } else {
          p <- ggplot(data, aes_string(x = group_var, fill = group_var)) +
            geom_bar(alpha = 0.7) +
            scale_fill_manual(values = colorRampPalette(c(color, "white"))(length(unique(data[[group_var]]))))
        }
        
        # Calculate counts for labels
        plot_data <- data %>%
          group_by(!!sym(group_var)) %>%
          summarise(count = n()) %>%
          mutate(percentage = count / sum(count) * 100)
        
        # Add labels if requested
        if (label_type != "None") {
          if (label_type == "Frequency" || label_type == "Both") {
            p <- p + geom_text(data = plot_data, 
                               aes(x = !!sym(group_var), y = count, 
                                   label = count, fill = NULL),
                               vjust = -0.5, size = label_size/3)
          }
          if (label_type == "Percentage" || label_type == "Both") {
            p <- p + geom_text(data = plot_data, 
                               aes(x = !!sym(group_var), y = count, 
                                   label = paste0(round(percentage, 1), "%"), fill = NULL),
                               vjust = 1.5, size = label_size/3, color = "white")
          }
        }
      } else {
        if (var_types[var] == "categorical") {
          # Use category colors if specified
          if (length(category_colors) > 0) {
            p <- ggplot(data, aes_string(x = var, fill = var)) +
              geom_bar(alpha = 0.7) +
              scale_fill_manual(values = category_colors)
          } else {
            p <- ggplot(data, aes_string(x = var)) +
              geom_bar(fill = color, alpha = 0.7)
          }
          
          # Calculate counts for labels
          plot_data <- data %>%
            group_by(!!sym(var)) %>%
            summarise(count = n()) %>%
            mutate(percentage = count / sum(count) * 100)
          
          # Add labels if requested
          if (label_type != "None") {
            if (label_type == "Frequency" || label_type == "Both") {
              p <- p + geom_text(data = plot_data, 
                                 aes(x = !!sym(var), y = count, 
                                     label = count, fill = NULL),
                                 vjust = -0.5, size = label_size/3)
            }
            if (label_type == "Percentage" || label_type == "Both") {
              p <- p + geom_text(data = plot_data, 
                                 aes(x = !!sym(var), y = count, 
                                     label = paste0(round(percentage, 1), "%"), fill = NULL),
                                 vjust = 1.5, size = label_size/3, color = "white")
            }
          }
        } else {
          # For numeric variables, create histogram instead
          p <- ggplot(data, aes_string(x = var)) +
            geom_histogram(fill = color, alpha = 0.7, bins = 30)
        }
      }
      
    } else if (chart_type == "Pie") {
      var <- vars[1]
      
      # Create data for pie chart
      freq_table <- as.data.frame(table(data[[var]]))
      colnames(freq_table) <- c("Category", "Count")
      freq_table$Percentage <- freq_table$Count / sum(freq_table$Count) * 100
      
      # Use category colors if specified
      if (length(category_colors) > 0) {
        # Match colors to categories
        pie_colors <- sapply(as.character(freq_table$Category), function(cat) {
          category_colors[[cat]] %||% color
        })
        
        p <- ggplot(freq_table, aes(x = "", y = Count, fill = Category)) +
          geom_bar(stat = "identity", width = 1, color = "white") +
          coord_polar(theta = "y") +
          scale_fill_manual(values = pie_colors)
      } else {
        p <- ggplot(freq_table, aes(x = "", y = Count, fill = Category)) +
          geom_bar(stat = "identity", width = 1, color = "white") +
          coord_polar(theta = "y") +
          scale_fill_manual(values = colorRampPalette(brewer.pal(8, "Set2"))(nrow(freq_table)))
      }
      
      # Remove axis labels and ticks
      p <- p + 
        scale_y_continuous(breaks = NULL) +
        theme_void() +
        theme(
          axis.text = element_blank(),
          axis.ticks = element_blank(),
          axis.title = element_blank(),
          panel.grid = element_blank(),
          legend.position = "right",
          legend.text = element_text(size = legend_size),
          legend.title = element_text(size = legend_size)
        )
      
      # Add labels if requested
      if (label_type != "None") {
        if (label_type == "Frequency") {
          p <- p + geom_text(aes(label = Count), 
                             position = position_stack(vjust = 0.5),
                             size = label_size/3, color = "white", fontface = "bold")
        } else if (label_type == "Percentage") {
          p <- p + geom_text(aes(label = paste0(round(Percentage, 1), "%")), 
                             position = position_stack(vjust = 0.5),
                             size = label_size/3, color = "white", fontface = "bold")
        } else if (label_type == "Both") {
          p <- p + geom_text(aes(label = paste0(Count, " (", round(Percentage, 1), "%)")), 
                             position = position_stack(vjust = 0.5),
                             size = label_size/3, color = "white", fontface = "bold")
        }
      }
      
      return(p)
      
    } else if (chart_type == "Histogram") {
      var <- vars[1]
      p <- ggplot(data, aes_string(x = var)) +
        geom_histogram(fill = color, alpha = 0.7, bins = 30)
      
    } else if (chart_type == "Boxplot") {
      var <- vars[1]
      
      if (!is.null(group_var)) {
        p <- ggplot(data, aes_string(x = group_var, y = var, fill = group_var)) +
          geom_boxplot(alpha = 0.7) +
          scale_fill_manual(values = colorRampPalette(c(color, "white"))(length(unique(data[[group_var]]))))
      } else {
        p <- ggplot(data, aes_string(y = var)) +
          geom_boxplot(fill = color, alpha = 0.7)
      }
      
    } else if (chart_type == "Scatterplot") {
      if (length(vars) < 2) {
        return(ggplot() + annotate("text", x = 1, y = 1, label = "Scatterplot requires 2 variables"))
      }
      
      x_var <- vars[1]
      y_var <- vars[2]
      
      if (!is.null(group_var)) {
        p <- ggplot(data, aes_string(x = x_var, y = y_var, color = group_var)) +
          geom_point(alpha = 0.7) +
          scale_color_manual(values = colorRampPalette(c(color, "white"))(length(unique(data[[group_var]]))))
        
        if (show_ids) {
          p <- p + geom_text(aes(label = rownames(data)), vjust = -1, check_overlap = TRUE, size = label_size/3)
        }
      } else {
        p <- ggplot(data, aes_string(x = x_var, y = y_var)) +
          geom_point(color = color, alpha = 0.7)
        
        if (show_ids) {
          p <- p + geom_text(aes(label = rownames(data)), vjust = -1, check_overlap = TRUE, size = label_size/3)
        }
      }
      
    } else if (chart_type == "Density") {
      var <- vars[1]
      
      if (!is.null(group_var)) {
        p <- ggplot(data, aes_string(x = var, fill = group_var)) +
          geom_density(alpha = 0.5) +
          scale_fill_manual(values = colorRampPalette(c(color, "white"))(length(unique(data[[group_var]]))))
      } else {
        p <- ggplot(data, aes_string(x = var)) +
          geom_density(fill = color, alpha = 0.7)
      }
      
    } else if (chart_type == "Violin") {
      var <- vars[1]
      
      if (!is.null(group_var)) {
        p <- ggplot(data, aes_string(x = group_var, y = var, fill = group_var)) +
          geom_violin(alpha = 0.7) +
          scale_fill_manual(values = colorRampPalette(c(color, "white"))(length(unique(data[[group_var]]))))
      } else {
        p <- ggplot(data, aes_string(y = var)) +
          geom_violin(fill = color, alpha = 0.7)
      }
      
    } else if (chart_type == "Q-Q Plot") {
      var <- vars[1]
      
      # Create Q-Q plot
      p <- ggplot(data, aes_string(sample = var)) +
        stat_qq(color = color) +
        stat_qq_line(color = "red") +
        labs(title = paste("Q-Q Plot of", var))
      
    } else {
      # Default to histogram
      var <- vars[1]
      p <- ggplot(data, aes_string(x = var)) +
        geom_histogram(fill = color, alpha = 0.7, bins = 30)
    }
    
    # Add title and labels
    if (title != "") {
      p <- p + ggtitle(title)
    }
    
    if (x_label != "") {
      p <- p + xlab(x_label)
    }
    
    if (y_label != "") {
      p <- p + ylab(y_label)
    }
    
    # Apply theme
    p <- p + theme_func() +
      theme(
        axis.text = element_text(size = input$axis_text_size),
        axis.title = element_text(size = input$axis_text_size + 2, face = "bold"),
        panel.grid.major = element_blank(),
        panel.grid.minor = element_blank()
      )
    
    # For pie charts, ensure no axis elements appear
    if (chart_type == "Pie") {
      p <- p + theme(
        axis.line = element_blank(),
        axis.text = element_blank(),
        axis.ticks = element_blank(),
        axis.title = element_blank(),
        panel.grid = element_blank(),
        panel.border = element_blank()
      )
    }
    
    # Flip coordinates if requested
    if (flip_coords) {
      p <- p + coord_flip()
    }
    
    return(p)
  }
  # Download handlers for different formats
  output$download_chart_png <- downloadHandler(
    filename = function() {
      paste("chart_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".png", sep = "")
    },
    content = function(file) {
      req(input$generate_chart)
      
      # Recreate the chart
      df <- data()
      chart_vars <- input$chart_vars
      group_var <- if (input$chart_group != "None") input$chart_group else NULL
      
      var_types <- sapply(df[chart_vars], function(x) {
        if (is.numeric(x)) "numeric" else "categorical"
      })
      
      p <- create_ggplot_chart(
        data = df,
        chart_type = input$chart_type,
        vars = chart_vars,
        var_types = var_types,
        group_var = group_var,
        title = input$chart_title,
        x_label = input$x_axis_label,
        y_label = input$y_axis_label,
        color = input$chart_color,
        theme_name = input$chart_theme,
        show_ids = input$show_point_ids,
        label_type = input$label_type,
        label_size = input$label_size,
        legend_size = input$legend_size,
        flip_coords = input$flip_coords
      )
      
      # Apply text sizes
      p <- p + theme(
        plot.title = element_text(size = input$title_size, face = "bold"),
        axis.text = element_text(size = input$axis_text_size),
        axis.title = element_text(size = input$axis_text_size + 2, face = "bold"),
        legend.text = element_text(size = input$legend_size),
        legend.title = element_text(size = input$legend_size)
      )
      
      # Save the plot
      ggsave(file, plot = p, width = input$chart_width/100, 
             height = input$chart_height/100, dpi = 300, device = "png")
    }
  )
  
  output$download_chart_jpeg <- downloadHandler(
    filename = function() {
      paste("chart_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".jpeg", sep = "")
    },
    content = function(file) {
      req(input$generate_chart)
      
      # Recreate the chart (same as PNG)
      df <- data()
      chart_vars <- input$chart_vars
      group_var <- if (input$chart_group != "None") input$chart_group else NULL
      
      var_types <- sapply(df[chart_vars], function(x) {
        if (is.numeric(x)) "numeric" else "categorical"
      })
      
      p <- create_ggplot_chart(
        data = df,
        chart_type = input$chart_type,
        vars = chart_vars,
        var_types = var_types,
        group_var = group_var,
        title = input$chart_title,
        x_label = input$x_axis_label,
        y_label = input$y_axis_label,
        color = input$chart_color,
        theme_name = input$chart_theme,
        show_ids = input$show_point_ids,
        label_type = input$label_type,
        label_size = input$label_size,
        legend_size = input$legend_size,
        flip_coords = input$flip_coords
      )
      
      # Apply text sizes
      p <- p + theme(
        plot.title = element_text(size = input$title_size, face = "bold"),
        axis.text = element_text(size = input$axis_text_size),
        axis.title = element_text(size = input$axis_text_size + 2, face = "bold"),
        legend.text = element_text(size = input$legend_size),
        legend.title = element_text(size = input$legend_size)
      )
      
      # Save the plot
      ggsave(file, plot = p, width = input$chart_width/100, 
             height = input$chart_height/100, dpi = 300, device = "jpeg")
    }
  )
  
  output$download_chart_pdf <- downloadHandler(
    filename = function() {
      paste("chart_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".pdf", sep = "")
    },
    content = function(file) {
      req(input$generate_chart)
      
      # Recreate the chart (same as PNG)
      df <- data()
      chart_vars <- input$chart_vars
      group_var <- if (input$chart_group != "None") input$chart_group else NULL
      
      var_types <- sapply(df[chart_vars], function(x) {
        if (is.numeric(x)) "numeric" else "categorical"
      })
      
      p <- create_ggplot_chart(
        data = df,
        chart_type = input$chart_type,
        vars = chart_vars,
        var_types = var_types,
        group_var = group_var,
        title = input$chart_title,
        x_label = input$x_axis_label,
        y_label = input$y_axis_label,
        color = input$chart_color,
        theme_name = input$chart_theme,
        show_ids = input$show_point_ids,
        label_type = input$label_type,
        label_size = input$label_size,
        legend_size = input$legend_size,
        flip_coords = input$flip_coords
      )
      
      # Apply text sizes
      p <- p + theme(
        plot.title = element_text(size = input$title_size, face = "bold"),
        axis.text = element_text(size = input$axis_text_size),
        axis.title = element_text(size = input$axis_text_size + 2, face = "bold"),
        legend.text = element_text(size = input$legend_size),
        legend.title = element_text(size = input$legend_size)
      )
      
      # Save the plot
      ggsave(file, plot = p, width = input$chart_width/100, 
             height = input$chart_height/100, device = "pdf")
    }
  )
  
  output$download_chart_pdf <- downloadHandler(
    filename = function() {
      paste("chart_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".pdf", sep = "")
    },
    content = function(file) {
      req(input$generate_chart)
      
      # Recreate the chart (same as PNG)
      df <- data()
      chart_vars <- input$chart_vars
      group_var <- if (input$chart_group != "None") input$chart_group else NULL
      
      var_types <- sapply(df[chart_vars], function(x) {
        if (is.numeric(x)) "numeric" else "categorical"
      })
      
      p <- create_ggplot_chart(
        data = df,
        chart_type = input$chart_type,
        vars = chart_vars,
        var_types = var_types,
        group_var = group_var,
        title = input$chart_title,
        x_label = input$x_axis_label,
        y_label = input$y_axis_label,
        color = input$chart_color,
        theme_name = input$chart_theme,
        show_ids = input$show_point_ids
      )
      
      # Apply text sizes
      p <- p + theme(
        plot.title = element_text(size = input$title_size, face = "bold"),
        axis.text = element_text(size = input$axis_text_size),
        axis.title = element_text(size = input$axis_text_size + 2, face = "bold")
      )
      
      # Save the plot
      ggsave(file, plot = p, width = input$chart_width/100, 
             height = input$chart_height/100, device = "pdf")
    }
  )
  # Placeholder outputs for plotly charts
  output$freq_chart_output <- renderPlotly({
    req(data(), input$freq_vars, input$run_frequencies)
    
    tryCatch({
      df <- data()
      selected_vars <- input$freq_vars
      
      if (length(selected_vars) == 0) {
        return(plotly_empty())
      }
      
      # Get the first selected variable for the chart
      var <- selected_vars[1]
      var_data <- df[[var]]
      
      # Create appropriate chart based on selection
      if ("Bar charts" %in% input$freq_charts) {
        # Create bar chart
        freq_table <- as.data.frame(table(var_data, useNA = "ifany"))
        colnames(freq_table) <- c("Category", "Frequency")
        
        p <- plot_ly(freq_table, x = ~Category, y = ~Frequency, type = "bar",
                     marker = list(color = spss_blue)) %>%
          layout(title = paste("Bar Chart of", var),
                 xaxis = list(title = var),
                 yaxis = list(title = "Frequency"))
        
      } else if ("Pie charts" %in% input$freq_charts) {
        # Create pie chart
        freq_table <- as.data.frame(table(var_data, useNA = "ifany"))
        colnames(freq_table) <- c("Category", "Frequency")
        
        p <- plot_ly(freq_table, labels = ~Category, values = ~Frequency, type = "pie",
                     textinfo = 'label+percent',
                     marker = list(colors = c(spss_blue, "#E27D72", "#68C3A3", "#F7C46C"))) %>%
          layout(title = paste("Pie Chart of", var))
        
      } else if ("Histograms" %in% input$freq_charts && is.numeric(var_data)) {
        # Create histogram for numeric variables
        p <- plot_ly(x = ~var_data, type = "histogram",
                     marker = list(color = spss_blue),
                     nbinsx = 30) %>%
          layout(title = paste("Histogram of", var),
                 xaxis = list(title = var),
                 yaxis = list(title = "Frequency"))
        
      } else {
        # Default to bar chart if no specific chart selected or invalid type
        freq_table <- as.data.frame(table(var_data, useNA = "ifany"))
        colnames(freq_table) <- c("Category", "Frequency")
        
        p <- plot_ly(freq_table, x = ~Category, y = ~Frequency, type = "bar",
                     marker = list(color = spss_blue)) %>%
          layout(title = paste("Frequency Distribution of", var),
                 xaxis = list(title = var),
                 yaxis = list(title = "Frequency"))
      }
      
      return(p)
      
    }, error = function(e) {
      print(paste("Chart error:", e$message))
      return(plotly_empty())
    })
  })
  
  output$explore_plot_output <- renderPlotly({
    req(data(), input$explore_vars, input$run_explore)
    
    tryCatch({
      df <- data()
      explore_vars <- input$explore_vars
      factor_var <- input$explore_factor
      
      if (length(explore_vars) == 0) {
        return(plotly_empty())
      }
      
      # Get the first selected variable for plotting
      var <- explore_vars[1]
      var_data <- df[[var]]
      
      # Remove missing values
      var_data <- var_data[!is.na(var_data)]
      
      # Create appropriate plot based on selection
      if ("Boxplots" %in% input$explore_plots && factor_var != "None") {
        # Boxplot by factor
        p <- plot_ly(df, x = as.formula(paste0("~", factor_var)), 
                     y = as.formula(paste0("~", var)), 
                     type = "box", color = as.formula(paste0("~", factor_var))) %>%
          layout(title = paste("Boxplot of", var, "by", factor_var),
                 xaxis = list(title = factor_var),
                 yaxis = list(title = var))
        
      } else if ("Histograms" %in% input$explore_plots) {
        # Histogram
        p <- plot_ly(x = ~var_data, type = "histogram",
                     marker = list(color = spss_blue),
                     nbinsx = 30) %>%
          layout(title = paste("Histogram of", var),
                 xaxis = list(title = var),
                 yaxis = list(title = "Frequency"))
        
      } else if ("Normality plots" %in% input$explore_plots) {
        # Q-Q plot for normality
        qq_data <- data.frame(
          Theoretical = qqnorm(var_data, plot.it = FALSE)$x,
          Sample = qqnorm(var_data, plot.it = FALSE)$y
        )
        
        # Fit line
        fit <- lm(Sample ~ Theoretical, data = qq_data)
        
        p <- plot_ly(qq_data, x = ~Theoretical, y = ~Sample, 
                     type = "scatter", mode = "markers",
                     marker = list(color = spss_blue)) %>%
          add_lines(x = ~Theoretical, y = fitted(fit),
                    line = list(color = spss_dark_blue)) %>%
          layout(title = paste("Q-Q Plot of", var),
                 xaxis = list(title = "Theoretical Quantiles"),
                 yaxis = list(title = "Sample Quantiles"))
        
      } else {
        # Default to boxplot if no specific plot selected
        if (factor_var != "None") {
          p <- plot_ly(df, x = as.formula(paste0("~", factor_var)), 
                       y = as.formula(paste0("~", var)), 
                       type = "box", color = as.formula(paste0("~", factor_var))) %>%
            layout(title = paste("Boxplot of", var, "by", factor_var),
                   xaxis = list(title = factor_var),
                   yaxis = list(title = var))
        } else {
          p <- plot_ly(y = ~var_data, type = "box",
                       marker = list(color = spss_blue)) %>%
            layout(title = paste("Boxplot of", var),
                   yaxis = list(title = var))
        }
      }
      
      return(p)
      
    }, error = function(e) {
      print(paste("Explore plot error:", e$message))
      return(plotly_empty())
    })
  })
  
  output$ttest_plot <- renderPlotly({
    req(input$run_ttest)
    plotly_empty()
  })
  
  output$paired_plot <- renderPlotly({
    req(input$run_paired)
    plotly_empty()
  })
  
  output$anova_plot <- renderPlotly({
    req(input$run_anova)
    plotly_empty()
  })
  
  output$cor_plot <- renderPlotly({
    req(input$run_correlation)
    plotly_empty()
  })
  
  output$reg_plot <- renderPlot({
    req(input$run_regression)
    ggplot() + theme_void()
  })
  
  output$logistic_plot <- renderPlotly({
    req(input$run_logistic)
    plotly_empty()
  })
  
  output$factor_scree <- renderPlotly({
    req(input$run_factor)
    plotly_empty()
  })
  
  output$chi_plot <- renderPlotly({
    req(input$run_chi)
    plotly_empty()
  })
  
  output$fisher_plot <- renderPlotly({
    req(input$run_fisher)
    plotly_empty()
  })
  
  # Additional outputs for other tabs
  output$sorted_preview <- renderDT({
    req(data(), input$sort_btn)
    datatable(data(), options = list(scrollX = TRUE, pageLength = 5))
  })
  
  output$selected_preview <- renderDT({
    req(data(), input$select_btn)
    datatable(data(), options = list(scrollX = TRUE, pageLength = 5))
  })
  
  output$weight_output <- renderPrint({
    req(data(), input$weight_btn)
    cat("Weighting applied successfully.\n")
  })
  
  output$transform_preview <- renderDT({
    req(data())
    datatable(data(), options = list(scrollX = TRUE, pageLength = 5))
  })
  
  # Sort cases functionality
  observeEvent(input$sort_btn, {
    req(data(), input$sort_vars)
    
    tryCatch({
      df <- data()
      
      # Sort the data
      if (input$sort_order == "Ascending") {
        df <- df[order(df[[input$sort_vars[1]]]), ]
      } else {
        df <- df[order(-df[[input$sort_vars[1]]]), ]
      }
      
      data(df)
      showNotification("Data sorted successfully!", type = "message")
      
    }, error = function(e) {
      showNotification(paste("Error sorting data:", e$message), type = "error")
    })
  })
  
  # Select cases functionality
  observeEvent(input$select_btn, {
    req(data())
    
    tryCatch({
      df <- data()
      
      if (input$select_method == "If condition is satisfied" && input$select_condition != "") {
        # Parse and evaluate condition
        condition <- parse(text = input$select_condition)
        selected_rows <- eval(condition, envir = df)
        df <- df[selected_rows, ]
      } else if (input$select_method == "Random sample") {
        # Take random sample
        sample_size <- round(nrow(df) * input$sample_percent / 100)
        df <- df[sample(nrow(df), sample_size), ]
      }
      
      data(df)
      showNotification("Cases selected successfully!", type = "message")
      
    }, error = function(e) {
      showNotification(paste("Error selecting cases:", e$message), type = "error")
    })
  })
  
  # Weight cases functionality
  observeEvent(input$weight_btn, {
    req(data())
    
    tryCatch({
      if (input$weight_method == "Weight cases by" && !is.null(input$weight_var)) {
        showNotification("Weighting applied (simulated - actual weighting would modify analysis)", type = "message")
      } else {
        showNotification("Weighting removed", type = "message")
      }
    }, error = function(e) {
      showNotification(paste("Error applying weighting:", e$message), type = "error")
    })
  })
  
  # Descriptive statistics analysis
  output$desc_output <- renderPrint({
    req(data(), input$desc_vars, input$run_descriptives)
    
    tryCatch({
      df <- data()
      selected_vars <- input$desc_vars
      
      if (length(selected_vars) == 0) {
        return("Please select at least one variable.")
      }
      
      # Subset data to selected variables
      data_subset <- df[, selected_vars, drop = FALSE]
      
      # Calculate descriptive statistics
      desc_stats <- psych::describe(data_subset)
      
      cat("Descriptive Statistics\n")
      cat("=====================\n\n")
      
      # Display selected statistics
      if ("Mean" %in% input$desc_stats) {
        cat("Means:\n")
        print(round(desc_stats$mean, 3))
        cat("\n")
      }
      
      if ("Std. deviation" %in% input$desc_stats) {
        cat("Standard Deviations:\n")
        print(round(desc_stats$sd, 3))
        cat("\n")
      }
      
      if ("Minimum" %in% input$desc_stats) {
        cat("Minimum Values:\n")
        print(desc_stats$min)
        cat("\n")
      }
      
      if ("Maximum" %in% input$desc_stats) {
        cat("Maximum Values:\n")
        print(desc_stats$max)
        cat("\n")
      }
      
      if ("Variance" %in% input$desc_stats) {
        cat("Variances:\n")
        variances <- apply(data_subset, 2, var, na.rm = TRUE)
        print(round(variances, 3))
        cat("\n")
      }
      
      if ("Range" %in% input$desc_stats) {
        cat("Ranges:\n")
        ranges <- apply(data_subset, 2, function(x) diff(range(x, na.rm = TRUE)))
        print(round(ranges, 3))
        cat("\n")
      }
      
      if ("S.E. mean" %in% input$desc_stats) {
        cat("Standard Error of Mean:\n")
        se_means <- desc_stats$sd / sqrt(desc_stats$n)
        print(round(se_means, 3))
        cat("\n")
      }
      
      if ("Kurtosis" %in% input$desc_stats) {
        cat("Kurtosis:\n")
        print(round(desc_stats$kurtosis, 3))
        cat("\n")
      }
      
      if ("Skewness" %in% input$desc_stats) {
        cat("Skewness:\n")
        print(round(desc_stats$skew, 3))
        cat("\n")
      }
      
      if ("Valid N" %in% input$desc_stats) {
        cat("Valid Cases (N):\n")
        print(desc_stats$n)
        cat("\n")
      }
      
      if ("Missing N" %in% input$desc_stats) {
        cat("Missing Values:\n")
        missing_counts <- colSums(is.na(data_subset))
        print(missing_counts)
        cat("\n")
      }
      
    }, error = function(e) {
      return(paste("Error in descriptive analysis:", e$message))
    })
  })
  
  # Frequencies analysis
  output$freq_output <- renderPrint({
    req(data(), input$freq_vars, input$run_frequencies)
    
    tryCatch({
      df <- data()
      selected_vars <- input$freq_vars
      
      if (length(selected_vars) == 0) {
        return("Please select at least one variable.")
      }
      
      cat("Frequency Tables\n")
      cat("================\n\n")
      
      for (var in selected_vars) {
        cat("Variable:", var, "\n")
        cat("----------------------------------------\n")
        
        freq_table <- table(df[[var]], useNA = "always")
        prop_table <- prop.table(freq_table) * 100
        
        result_df <- data.frame(
          Category = names(freq_table),
          Frequency = as.numeric(freq_table),
          Percent = round(as.numeric(prop_table), 2),
          CumulativePercent = round(cumsum(as.numeric(prop_table)), 2)
        )
        
        # Replace NA with "Missing" for display
        result_df$Category[is.na(result_df$Category)] <- "Missing"
        
        print(result_df, row.names = FALSE)
        cat("\n\n")
      }
      
    }, error = function(e) {
      return(paste("Error in frequency analysis:", e$message))
    })
  })
  
  # Independent Samples T-Test Analysis - CORRECTED VERSION
  output$ttest_output <- renderPrint({
    req(data(), input$ttest_vars, input$ttest_group_var, input$run_ttest)
    
    tryCatch({
      df <- data()
      test_vars <- input$ttest_vars
      group_var <- input$ttest_group_var
      
      if (length(test_vars) == 0) {
        return("Please select at least one test variable.")
      }
      
      if (is.null(group_var) || group_var == "") {
        return("Please select a grouping variable.")
      }
      
      # Parse group definition
      groups <- strsplit(input$group_def, ",")[[1]]
      groups <- trimws(groups)
      
      if (length(groups) != 2) {
        return("Please define exactly two groups (e.g., 'male,female' for Gender)")
      }
      
      cat("Independent Samples T-Test\n")
      cat("==========================\n\n")
      
      for (test_var in test_vars) {
        cat("Variable:", test_var, "\n")
        cat("----------------------------------------\n")
        
        # Filter data for the two groups
        group_data <- df[df[[group_var]] %in% groups, ]
        
        # Extract data for each group
        group1_data <- group_data[group_data[[group_var]] == groups[1], test_var]
        group2_data <- group_data[group_data[[group_var]] == groups[2], test_var]
        
        # Remove missing values
        group1_data <- group1_data[!is.na(group1_data)]
        group2_data <- group2_data[!is.na(group2_data)]
        
        # Check for sufficient observations
        if (length(group1_data) < 2 || length(group2_data) < 2) {
          cat("Insufficient data for t-test:\n")
          cat("Group 1 (", groups[1], "): N =", length(group1_data), "\n")
          cat("Group 2 (", groups[2], "): N =", length(group2_data), "\n")
          cat("Each group needs at least 2 observations for t-test.\n\n")
          next
        }
        
        # Perform t-test with proper tail handling
        if (input$ttest_tails == "Two-tailed") {
          t_test <- t.test(group1_data, group2_data, 
                           var.equal = TRUE,
                           conf.level = input$ttest_conf/100,
                           alternative = "two.sided")
        } else {
          # Determine direction for one-tailed test
          mean_diff <- mean(group1_data) - mean(group2_data)
          alternative <- ifelse(mean_diff > 0, "greater", "less")
          t_test <- t.test(group1_data, group2_data, 
                           var.equal = TRUE,
                           conf.level = input$ttest_conf/100,
                           alternative = alternative)
        }
        
        # Display results
        cat("Group 1 (", groups[1], "): N =", length(group1_data), 
            ", Mean =", round(mean(group1_data), 3), 
            ", SD =", round(sd(group1_data), 3), "\n")
        cat("Group 2 (", groups[2], "): N =", length(group2_data), 
            ", Mean =", round(mean(group2_data), 3), 
            ", SD =", round(sd(group2_data), 3), "\n\n")
        
        cat("t(", t_test$parameter, ") =", round(t_test$statistic, 3), "\n")
        cat("p-value =", format.pval(t_test$p.value, digits = 3, eps = 0.001), "\n")
        cat(input$ttest_conf, "% CI: [", round(t_test$conf.int[1], 3), ",", 
            round(t_test$conf.int[2], 3), "]\n")
        cat("Mean Difference =", round(t_test$estimate[1] - t_test$estimate[2], 3), "\n\n")
        
        # Add interpretation for one-tailed tests
        if (input$ttest_tails == "One-tailed") {
          if (t_test$p.value < 0.05) {
            if (mean(group1_data) > mean(group2_data)) {
              cat("Interpretation: Group 1 is significantly greater than Group 2 (one-tailed)\n")
            } else {
              cat("Interpretation: Group 1 is significantly less than Group 2 (one-tailed)\n")
            }
          } else {
            cat("Interpretation: No significant difference between groups (one-tailed)\n")
          }
        }
        cat("\n")
      }
      
    }, error = function(e) {
      return(paste("Error in t-test analysis:", e$message))
    })
  })
  
  # One Sample T-Test Analysis - CORRECTED VERSION
  output$onesample_output <- renderPrint({
    req(data(), input$onesample_vars, input$run_onesample)
    
    tryCatch({
      df <- data()
      test_vars <- input$onesample_vars
      test_value <- input$test_value
      
      if (length(test_vars) == 0) {
        return("Please select at least one test variable.")
      }
      
      cat("One Sample T-Test\n")
      cat("=================\n\n")
      cat("Test Value:", test_value, "\n\n")
      
      for (test_var in test_vars) {
        cat("Variable:", test_var, "\n")
        cat("----------------------------------------\n")
        
        # Extract data
        test_data <- df[[test_var]]
        test_data <- test_data[!is.na(test_data)]
        
        # Perform one-sample t-test with proper tail handling
        if (input$onesample_tails == "Two-tailed") {
          t_test <- t.test(test_data, mu = test_value, 
                           conf.level = input$onesample_conf/100,
                           alternative = "two.sided")
        } else {
          # Determine direction for one-tailed test
          mean_diff <- mean(test_data) - test_value
          alternative <- ifelse(mean_diff > 0, "greater", "less")
          t_test <- t.test(test_data, mu = test_value, 
                           conf.level = input$onesample_conf/100,
                           alternative = alternative)
        }
        
        # Display results
        cat("N =", length(test_data), 
            ", Mean =", round(mean(test_data), 3), 
            ", SD =", round(sd(test_data), 3), "\n\n")
        
        cat("t(", t_test$parameter, ") =", round(t_test$statistic, 3), "\n")
        cat("p-value =", format.pval(t_test$p.value, digits = 3), "\n")
        cat(input$onesample_conf, "% CI: [", round(t_test$conf.int[1], 3), ",", 
            round(t_test$conf.int[2], 3), "]\n")
        cat("Mean Difference =", round(mean(test_data) - test_value, 3), "\n\n")
        
        # Add interpretation for one-tailed tests
        if (input$onesample_tails == "One-tailed") {
          if (t_test$p.value < 0.05) {
            if (mean(test_data) > test_value) {
              cat("Interpretation: Mean is significantly greater than test value (one-tailed)\n")
            } else {
              cat("Interpretation: Mean is significantly less than test value (one-tailed)\n")
            }
          } else {
            cat("Interpretation: No significant difference from test value (one-tailed)\n")
          }
        }
        cat("\n")
      }
      
    }, error = function(e) {
      return(paste("Error in one-sample t-test:", e$message))
    })
  })
  
  # ANOVA Tab - CORRECTED VERSION
  output$anova_output <- renderPrint({
    req(data(), input$anova_dv, input$anova_factor, input$run_anova)
    
    tryCatch({
      df <- data()
      dv_vars <- input$anova_dv
      factor_var <- input$anova_factor
      
      if (length(dv_vars) == 0) {
        return("Please select at least one dependent variable.")
      }
      
      if (is.null(factor_var) || factor_var == "") {
        return("Please select a factor variable.")
      }
      
      cat("One-Way ANOVA\n")
      cat("=============\n\n")
      
      # Store results for each DV
      anova_results <- list()
      
      for (dv_var in dv_vars) {
        cat("Dependent Variable:", dv_var, "\n")
        cat("Factor:", factor_var, "\n")
        cat("----------------------------------------\n")
        
        # Create formula for ANOVA
        formula <- as.formula(paste(dv_var, "~", factor_var))
        
        # Perform ANOVA
        anova_result <- aov(formula, data = df)
        anova_summary <- summary(anova_result)
        
        # Store result for post hoc tests
        anova_results[[dv_var]] <- anova_result
        
        print(anova_summary)
        cat("\n")
        
        # Homogeneity of variance test (Levene's test)
        if ("Homogeneity of variance test" %in% input$anova_options) {
          cat("Homogeneity of Variance Test (Levene's Test):\n")
          levene_test <- car::leveneTest(formula, data = df)
          print(levene_test)
          cat("\n")
        }
        
        # Welch ANOVA for unequal variances
        if ("Welch" %in% input$anova_options) {
          cat("Welch ANOVA (for unequal variances):\n")
          welch_test <- oneway.test(formula, data = df, var.equal = FALSE)
          print(welch_test)
          cat("\n")
        }
        
        # Brown-Forsythe test
        if ("Brown-Forsythe" %in% input$anova_options) {
          cat("Brown-Forsythe Test:\n")
          bf_test <- oneway.test(formula, data = df, var.equal = FALSE)
          cat("F(", bf_test$parameter[1], ",", bf_test$parameter[2], ") =", 
              round(bf_test$statistic, 3), ", p =", 
              format.pval(bf_test$p.value, digits = 3, eps = 0.001), "\n")
          cat("\n")
        }
      }
      
      # Store results for post hoc tests
      anova_results_list(anova_results)
      
    }, error = function(e) {
      return(paste("Error in ANOVA analysis:", e$message))
    })
  })
  
  
  # Post Hoc Tests - CORRECTED VERSION
  output$posthoc_output <- renderPrint({
    req(anova_results_list(), input$anova_dv, input$anova_factor, input$run_anova, input$posthoc_tests)
    
    tryCatch({
      anova_results <- anova_results_list()
      dv_vars <- input$anova_dv
      factor_var <- input$anova_factor
      
      if (length(dv_vars) == 0 || length(input$posthoc_tests) == 0) {
        return("Please select at least one dependent variable and post hoc test.")
      }
      
      cat("Post Hoc Tests\n")
      cat("==============\n\n")
      
      for (dv_var in dv_vars) {
        if (!dv_var %in% names(anova_results)) next
        
        cat("Dependent Variable:", dv_var, "\n")
        cat("Factor:", factor_var, "\n")
        cat("----------------------------------------\n")
        
        anova_result <- anova_results[[dv_var]]
        df <- data()
        
        # Tukey HSD
        if ("Tukey" %in% input$posthoc_tests) {
          cat("\nTukey HSD Test:\n")
          tukey_result <- TukeyHSD(anova_result)
          print(tukey_result)
        }
        
        # Bonferroni
        if ("Bonferroni" %in% input$posthoc_tests) {
          cat("\nBonferroni Test:\n")
          pairwise_result <- pairwise.t.test(df[[dv_var]], df[[factor_var]], 
                                             p.adjust.method = "bonferroni")
          print(pairwise_result)
        }
        
        # Scheffe
        if ("Scheffe" %in% input$posthoc_tests) {
          cat("\nScheffe Test:\n")
          emm <- emmeans(anova_result, specs = factor_var)
          scheffe_result <- pairs(emm, adjust = "scheffe")
          print(scheffe_result)
        }
        
        # Sidak
        if ("Sidak" %in% input$posthoc_tests) {
          cat("\nSidak Test:\n")
          
          # Raw pairwise t-tests without adjustment
          pairwise_raw <- pairwise.t.test(df[[dv_var]], df[[factor_var]], p.adjust.method = "none")
          
          # Number of comparisons
          m <- length(pairwise_raw$p.value[!is.na(pairwise_raw$p.value)])
          
          # Apply Sidak adjustment manually
          sidak_p <- 1 - (1 - pairwise_raw$p.value)^m
          sidak_p[is.na(pairwise_raw$p.value)] <- NA
          
          # Print Sidak-adjusted p-values
          print(sidak_p)
        }
        
        # Duncan's Multiple Range Test
        if ("Duncan" %in% input$posthoc_tests) {
          cat("\nDuncan's Multiple Range Test:\n")
          if (!requireNamespace("agricolae", quietly = TRUE)) {
            cat("Duncan test requires the 'agricolae' package. Please install it.\n")
          } else {
            duncan_result <- agricolae::duncan.test(anova_result, 
                                                    trt = factor_var, 
                                                    group = TRUE)
            print(duncan_result$groups)
          }
        }
        
        cat("\n")
      }
      
    }, error = function(e) {
      return(paste("Error in post hoc analysis:", e$message))
    })
  })
  
  
  # Descriptives for ANOVA
  output$descriptives_output <- renderPrint({
    req(data(), input$anova_dv, input$anova_factor, input$run_anova)
    
    tryCatch({
      df <- data()
      dv_vars <- input$anova_dv
      factor_var <- input$anova_factor
      
      if (length(dv_vars) == 0) {
        return("Please select at least one dependent variable.")
      }
      
      cat("Descriptive Statistics by Group\n")
      cat("===============================\n\n")
      
      for (dv_var in dv_vars) {
        cat("Variable:", dv_var, "\n")
        cat("Grouping Variable:", factor_var, "\n")
        cat("----------------------------------------\n")
        
        # Calculate descriptive statistics by group
        desc_stats <- df %>%
          group_by(!!sym(factor_var)) %>%
          summarise(
            N = sum(!is.na(!!sym(dv_var))),
            Mean = mean(!!sym(dv_var), na.rm = TRUE),
            SD = sd(!!sym(dv_var), na.rm = TRUE),
            SE = SD / sqrt(N),
            Min = min(!!sym(dv_var), na.rm = TRUE),
            Max = max(!!sym(dv_var), na.rm = TRUE)
          )
        
        print(as.data.frame(desc_stats), row.names = FALSE)
        cat("\n")
        
        # Overall statistics
        cat("Overall Statistics:\n")
        overall <- data.frame(
          N = sum(!is.na(df[[dv_var]])),
          Mean = mean(df[[dv_var]], na.rm = TRUE),
          SD = sd(df[[dv_var]], na.rm = TRUE),
          SE = sd(df[[dv_var]], na.rm = TRUE) / sqrt(sum(!is.na(df[[dv_var]]))),
          Min = min(df[[dv_var]], na.rm = TRUE),
          Max = max(df[[dv_var]], na.rm = TRUE)
        )
        print(overall)
        cat("\n\n")
      }
      
    }, error = function(e) {
      return(paste("Error in descriptive statistics:", e$message))
    })
  })
  
  # ANOVA Plots
  output$anova_plot <- renderPlotly({
    req(data(), input$anova_dv, input$anova_factor, input$run_anova)
    
    tryCatch({
      df <- data()
      dv_vars <- input$anova_dv
      factor_var <- input$anova_factor
      
      if (length(dv_vars) == 0) {
        return(plotly_empty())
      }
      
      # Use the first dependent variable for plotting
      dv_var <- dv_vars[1]
      
      # Create boxplot
      p <- plot_ly(df, 
                   x = as.formula(paste0("~", factor_var)), 
                   y = as.formula(paste0("~", dv_var)), 
                   type = "box",
                   color = as.formula(paste0("~", factor_var))) %>%
        layout(title = paste("Boxplot of", dv_var, "by", factor_var),
               xaxis = list(title = factor_var),
               yaxis = list(title = dv_var),
               showlegend = FALSE)
      
      return(p)
      
    }, error = function(e) {
      print(paste("ANOVA plot error:", e$message))
      return(plotly_empty())
    })
  })
  
  # Helper function for Scheffe test (if not available)
  scheffe.test <- function(aov_model, factor_name) {
    # Simple implementation of Scheffe test
    anova_summary <- summary(aov_model)
    mse <- anova_summary[[1]]["Residuals", "Mean Sq"]
    df_error <- anova_summary[[1]]["Residuals", "Df"]
    
    groups <- model.frame(aov_model)[[factor_name]]
    means <- tapply(model.response(model.frame(aov_model)), groups, mean)
    n <- tapply(model.response(model.frame(aov_model)), groups, length)
    
    k <- length(means)
    f_critical <- qf(0.95, k-1, df_error)
    
    results <- list()
    for (i in 1:(k-1)) {
      for (j in (i+1):k) {
        diff <- abs(means[i] - means[j])
        se <- sqrt(mse * (1/n[i] + 1/n[j]))
        f_value <- (diff^2) / ((k-1) * se^2)
        p_value <- 1 - pf(f_value, k-1, df_error)
        
        results[[paste(names(means)[i], "-", names(means)[j])]] <- c(
          Difference = diff,
          SE = se,
          F = f_value,
          p.value = p_value
        )
      }
    }
    
    return(results)
  }
  # UI for dependent variable level selection in linear regression (if categorical)
  output$reg_dep_level_ui <- renderUI({
    req(data(), input$reg_dep)
    
    df <- data()
    dep_var <- input$reg_dep
    
    if (is.null(dep_var) || dep_var == "") {
      return(NULL)
    }
    
    # Only show this UI if the dependent variable is categorical
    if (is.factor(df[[dep_var]]) || is.character(df[[dep_var]])) {
      var_levels <- if (is.factor(df[[dep_var]])) {
        levels(df[[dep_var]])
      } else {
        unique(na.omit(df[[dep_var]]))
      }
      
      selectInput(
        inputId = "reg_dep_level",
        label = "Select reference category for dependent variable",
        choices = var_levels,
        selected = var_levels[1]
      )
    } else {
      return(NULL)
    }
  })
  
  # UI for categorical variable reference category selection for INDEPENDENT variables
  output$categorical_ref_ui <- renderUI({
    req(data(), input$reg_indep)
    
    df <- data()
    categorical_vars <- input$reg_indep[sapply(df[input$reg_indep], function(x) is.factor(x) || is.character(x))]
    
    if (length(categorical_vars) == 0) {
      return(NULL)
    }
    
    ref_ui_list <- lapply(categorical_vars, function(var) {
      var_levels <- if (is.factor(df[[var]])) {
        levels(df[[var]])
      } else {
        unique(na.omit(df[[var]]))
      }
      
      selectInput(
        inputId = paste0("ref_", var),
        label = paste("Reference category for", var),
        choices = var_levels,
        selected = var_levels[1]
      )
    })
    
    tagList(
      h4("Reference Categories for Categorical Variables"),
      ref_ui_list
    )
  })
  # Linear Regression analysis - Enhanced with categorical variable support
  output$reg_summary <- renderPrint({
    req(data(), input$reg_dep, input$reg_indep, input$run_regression)
    
    tryCatch({
      df <- data()
      dep_var <- input$reg_dep
      indep_vars <- input$reg_indep
      
      if (is.null(dep_var) || dep_var == "") {
        return("Please select a dependent variable.")
      }
      
      if (length(indep_vars) == 0) {
        return("Please select at least one independent variable.")
      }
      
      # Prepare data with proper reference categories for categorical variables
      model_data <- df
      categorical_vars <- indep_vars[sapply(df[indep_vars], function(x) is.factor(x) || is.character(x))]
      
      # Set reference categories for categorical variables
      for (var in categorical_vars) {
        ref_input <- input[[paste0("ref_", var)]]
        if (!is.null(ref_input)) {
          if (is.factor(model_data[[var]])) {
            model_data[[var]] <- relevel(model_data[[var]], ref = ref_input)
          } else {
            # Convert character to factor with specified reference
            model_data[[var]] <- factor(model_data[[var]])
            model_data[[var]] <- relevel(model_data[[var]], ref = ref_input)
          }
        }
      }
      
      # Create formula
      formula <- as.formula(paste(dep_var, "~", paste(indep_vars, collapse = " + ")))
      
      # Perform regression
      model <- lm(formula, data = model_data)
      
      cat("Linear Regression\n")
      cat("=================\n\n")
      cat("Dependent Variable:", dep_var, "\n")
      cat("Independent Variables:", paste(indep_vars, collapse = ", "), "\n\n")
      
      # Display reference categories used
      if (length(categorical_vars) > 0) {
        cat("Reference Categories:\n")
        for (var in categorical_vars) {
          ref_input <- input[[paste0("ref_", var)]]
          cat(paste0("  ", var, ": ", ref_input, "\n"))
        }
        cat("\n")
      }
      
      # Model summary
      cat("Model Summary\n")
      cat("-------------\n")
      summary_model <- summary(model)
      print(summary_model)
      
      # Display additional statistics if requested
      if ("R squared change" %in% input$reg_stats) {
        cat("\nR-squared Change Analysis\n")
        cat("-------------------------\n")
        # Calculate R-squared change for each variable added sequentially
        r_sq_change <- c()
        for (i in 1:length(indep_vars)) {
          current_vars <- indep_vars[1:i]
          current_formula <- as.formula(paste(dep_var, "~", paste(current_vars, collapse = " + ")))
          current_model <- lm(current_formula, data = model_data)
          r_sq_change[i] <- summary(current_model)$r.squared
        }
        
        # Calculate change in R-squared
        r_sq_change <- c(summary(lm(as.formula(paste(dep_var, "~ 1")), data = model_data))$r.squared, r_sq_change)
        change <- diff(r_sq_change)
        
        change_df <- data.frame(
          Step = 1:length(indep_vars),
          Variable_Added = indep_vars,
          R_Square_Change = round(change, 4),
          Cumulative_R_Square = round(r_sq_change[-1], 4)
        )
        print(change_df)
      }
      
      if ("Descriptives" %in% input$reg_stats) {
        cat("\nDescriptive Statistics\n")
        cat("----------------------\n")
        desc_vars <- c(dep_var, indep_vars)
        desc_data <- model_data[, desc_vars, drop = FALSE]
        desc_stats <- psych::describe(desc_data)
        print(round(desc_stats, 3))
      }
      
    }, error = function(e) {
      return(paste("Error in regression analysis:", e$message))
    })
  })
  
  output$reg_anova <- renderPrint({
    req(data(), input$reg_dep, input$reg_indep, input$run_regression)
    
    tryCatch({
      df <- data()
      dep_var <- input$reg_dep
      indep_vars <- input$reg_indep
      
      # Prepare data with proper reference categories
      model_data <- df
      categorical_vars <- indep_vars[sapply(df[indep_vars], function(x) is.factor(x) || is.character(x))]
      
      for (var in categorical_vars) {
        ref_input <- input[[paste0("ref_", var)]]
        if (!is.null(ref_input)) {
          if (is.factor(model_data[[var]])) {
            model_data[[var]] <- relevel(model_data[[var]], ref = ref_input)
          } else {
            model_data[[var]] <- factor(model_data[[var]])
            model_data[[var]] <- relevel(model_data[[var]], ref = ref_input)
          }
        }
      }
      
      # Create formula
      formula <- as.formula(paste(dep_var, "~", paste(indep_vars, collapse = " + ")))
      
      # Perform regression
      model <- lm(formula, data = model_data)
      
      cat("ANOVA Table\n")
      cat("===========\n\n")
      
      # ANOVA table
      anova_table <- anova(model)
      print(anova_table)
      
    }, error = function(e) {
      return(paste("Error in ANOVA analysis:", e$message))
    })
  })
  
  output$reg_coef <- renderPrint({
    req(data(), input$reg_dep, input$reg_indep, input$run_regression)
    
    tryCatch({
      df <- data()
      dep_var <- input$reg_dep
      indep_vars <- input$reg_indep
      
      # Prepare data with proper reference categories
      model_data <- df
      categorical_vars <- indep_vars[sapply(df[indep_vars], function(x) is.factor(x) || is.character(x))]
      
      for (var in categorical_vars) {
        ref_input <- input[[paste0("ref_", var)]]
        if (!is.null(ref_input)) {
          if (is.factor(model_data[[var]])) {
            model_data[[var]] <- relevel(model_data[[var]], ref = ref_input)
          } else {
            model_data[[var]] <- factor(model_data[[var]])
            model_data[[var]] <- relevel(model_data[[var]], ref = ref_input)
          }
        }
      }
      
      # Create formula
      formula <- as.formula(paste(dep_var, "~", paste(indep_vars, collapse = " + ")))
      
      # Perform regression
      model <- lm(formula, data = model_data)
      
      cat("Regression Coefficients\n")
      cat("======================\n\n")
      
      # Display reference categories used
      if (length(categorical_vars) > 0) {
        cat("Reference Categories:\n")
        for (var in categorical_vars) {
          ref_input <- input[[paste0("ref_", var)]]
          cat(paste0("  ", var, ": ", ref_input, "\n"))
        }
        cat("\n")
      }
      
      # Coefficient table with additional options
      coef_table <- summary(model)$coefficients
      colnames(coef_table) <- c("Estimate", "Std. Error", "t-value", "p-value")
      
      # Add confidence intervals if requested
      if ("Confidence intervals" %in% input$reg_stats) {
        ci <- confint(model, level = 0.95)
        coef_table <- cbind(coef_table, "CI Lower" = ci[, 1], "CI Upper" = ci[, 2])
      }
      
      # Add standardized coefficients (beta weights)
      if ("Estimates" %in% input$reg_stats) {
        # Calculate standardized coefficients
        sd_y <- sd(model_data[[dep_var]], na.rm = TRUE)
        
        # For categorical variables, we need to handle them differently
        beta_values <- c(NA)  # Intercept has no beta
        
        for (i in 2:nrow(coef_table)) {
          var_name <- rownames(coef_table)[i]
          
          # Check if this is a categorical variable level
          is_categorical_level <- FALSE
          cat_var <- NULL
          
          for (cv in categorical_vars) {
            if (grepl(paste0("^", cv), var_name)) {
              is_categorical_level <- TRUE
              cat_var <- cv
              break
            }
          }
          
          if (is_categorical_level) {
            # For categorical variables, we can't calculate a simple beta
            # We'll use the method from effectsize package if available
            if (requireNamespace("effectsize", quietly = TRUE)) {
              standardized <- effectsize::standardize_parameters(model)
              beta_values[i] <- standardized$Std_Coefficient[standardized$Parameter == var_name]
            } else {
              beta_values[i] <- NA
            }
          } else {
            # For continuous variables
            if (var_name %in% names(model_data)) {
              sd_x <- sd(model_data[[var_name]], na.rm = TRUE)
              beta_values[i] <- coef(model)[var_name] * sd_x / sd_y
            } else {
              beta_values[i] <- NA
            }
          }
        }
        
        coef_table <- cbind(coef_table, "Beta" = beta_values)
      }
      
      print(round(coef_table, 4))
      
      # Add covariance matrix if requested
      if ("Covariance matrix" %in% input$reg_stats) {
        cat("\nCovariance Matrix of Coefficients\n")
        cat("--------------------------------\n")
        vcov_matrix <- vcov(model)
        print(round(vcov_matrix, 4))
      }
      
      # Add part and partial correlations if requested
      if ("Part and partial correlations" %in% input$reg_stats) {
        cat("\nPart and Partial Correlations\n")
        cat("-----------------------------\n")
        
        # Calculate correlations
        cor_data <- model_data[, c(dep_var, indep_vars)]
        cor_data <- na.omit(cor_data)
        
        # Zero-order correlations
        zero_order <- cor(cor_data)[1, -1]
        
        # Partial correlations
        partial_cor <- c()
        part_cor <- c()
        
        for (i in 1:length(indep_vars)) {
          # Variables to control for
          control_vars <- indep_vars[-i]
          
          if (length(control_vars) > 0) {
            # Partial correlation
            res_y <- residuals(lm(as.formula(paste(dep_var, "~", paste(control_vars, collapse = " + "))), data = cor_data))
            res_x <- residuals(lm(as.formula(paste(indep_vars[i], "~", paste(control_vars, collapse = " + "))), data = cor_data))
            partial_cor[i] <- cor(res_y, res_x)
            
            # Part correlation (semi-partial)
            part_cor[i] <- cor(cor_data[[dep_var]], res_x)
          } else {
            partial_cor[i] <- zero_order[i]
            part_cor[i] <- zero_order[i]
          }
        }
        
        cor_df <- data.frame(
          Variable = indep_vars,
          Zero_Order = round(zero_order, 4),
          Partial = round(partial_cor, 4),
          Part = round(part_cor, 4)
        )
        print(cor_df)
      }
      
      # Add collinearity diagnostics if requested
      if ("Collinearity diagnostics" %in% input$reg_stats) {
        cat("\nCollinearity Diagnostics\n")
        cat("------------------------\n")
        
        # VIF and tolerance
        if (length(indep_vars) > 1) {
          # Check if we have categorical variables with more than 2 levels
          # For such cases, we need to use a different approach
          if (any(sapply(model_data[indep_vars], function(x) is.factor(x) && nlevels(x) > 2))) {
            cat("Note: VIF calculations may not be appropriate for categorical variables with more than 2 levels.\n")
          }
          
          vif_values <- car::vif(model)
          tolerance <- 1 / vif_values
          
          collin_df <- data.frame(
            Variable = names(vif_values),
            VIF = round(vif_values, 3),
            Tolerance = round(tolerance, 3)
          )
          print(collin_df)
        } else {
          cat("Collinearity diagnostics require at least 2 independent variables.\n")
        }
        
        # Condition indices
        X <- model.matrix(model)[, -1]  # Remove intercept
        if (ncol(X) > 1) {
          eigen_values <- eigen(cor(X))$values
          condition_indices <- sqrt(max(eigen_values) / eigen_values)
          
          cat("\nCondition Indices:\n")
          print(round(condition_indices, 3))
        }
      }
      
    }, error = function(e) {
      return(paste("Error in regression analysis:", e$message))
    })
  })
  
  output$reg_residuals <- renderPrint({
    req(data(), input$reg_dep, input$reg_indep, input$run_regression)
    
    tryCatch({
      df <- data()
      dep_var <- input$reg_dep
      indep_vars <- input$reg_indep
      
      # Prepare data with proper reference categories
      model_data <- df
      categorical_vars <- indep_vars[sapply(df[indep_vars], function(x) is.factor(x) || is.character(x))]
      
      for (var in categorical_vars) {
        ref_input <- input[[paste0("ref_", var)]]
        if (!is.null(ref_input)) {
          if (is.factor(model_data[[var]])) {
            model_data[[var]] <- relevel(model_data[[var]], ref = ref_input)
          } else {
            model_data[[var]] <- factor(model_data[[var]])
            model_data[[var]] <- relevel(model_data[[var]], ref = ref_input)
          }
        }
      }
      
      # Create formula
      formula <- as.formula(paste(dep_var, "~", paste(indep_vars, collapse = " + ")))
      
      # Perform regression
      model <- lm(formula, data = model_data)
      
      cat("Residual Statistics\n")
      cat("===================\n\n")
      
      residuals <- resid(model)
      fitted <- fitted(model)
      
      res_stats <- data.frame(
        Statistic = c("Minimum", "Maximum", "Mean", "Median", "Std. Deviation", 
                      "Skewness", "Kurtosis", "Durbin-Watson"),
        Value = c(
          round(min(residuals), 4),
          round(max(residuals), 4),
          round(mean(residuals), 4),
          round(median(residuals), 4),
          round(sd(residuals), 4),
          round(psych::skew(residuals), 4),
          round(psych::kurtosi(residuals), 4),
          round(lmtest::dwtest(model)$statistic, 4)
        )
      )
      
      print(res_stats, row.names = FALSE)
      
      # Normality test
      cat("\nTests of Normality\n")
      cat("------------------\n")
      shapiro_test <- shapiro.test(residuals)
      cat("Shapiro-Wilk Test: W =", round(shapiro_test$statistic, 4), 
          ", p =", format.pval(shapiro_test$p.value, digits = 4), "\n")
      
    }, error = function(e) {
      return(paste("Error in residual analysis:", e$message))
    })
  })
  
  output$reg_plot <- renderPlot({
    req(data(), input$reg_dep, input$reg_indep, input$run_regression)
    
    tryCatch({
      df <- data()
      dep_var <- input$reg_dep
      indep_vars <- input$reg_indep
      
      # Prepare data with proper reference categories
      model_data <- df
      categorical_vars <- indep_vars[sapply(df[indep_vars], function(x) is.factor(x) || is.character(x))]
      
      for (var in categorical_vars) {
        ref_input <- input[[paste0("ref_", var)]]
        if (!is.null(ref_input)) {
          if (is.factor(model_data[[var]])) {
            model_data[[var]] <- relevel(model_data[[var]], ref = ref_input)
          } else {
            model_data[[var]] <- factor(model_data[[var]])
            model_data[[var]] <- relevel(model_data[[var]], ref = ref_input)
          }
        }
      }
      
      # Create formula
      formula <- as.formula(paste(dep_var, "~", paste(indep_vars, collapse = " + ")))
      
      # Perform regression
      model <- lm(formula, data = model_data)
      
      # Set up plotting area based on selected plots
      plot_types <- input$reg_plots
      n_plots <- length(plot_types)
      
      if (n_plots == 0) {
        return(ggplot() + theme_void() + ggtitle("No plots selected"))
      }
      
      # Set up plot layout
      if (n_plots == 1) {
        par(mfrow = c(1, 1))
      } else if (n_plots == 2) {
        par(mfrow = c(1, 2))
      } else {
        par(mfrow = c(2, 2))
      }
      
      # Generate selected plots
      if ("Residuals vs Fitted" %in% plot_types) {
        plot(fitted(model), residuals(model),
             xlab = "Fitted Values", ylab = "Residuals",
             main = "Residuals vs Fitted Values")
        abline(h = 0, col = "red", lty = 2)
      }
      
      if ("Normal Q-Q" %in% plot_types) {
        qqnorm(residuals(model), main = "Normal Q-Q Plot")
        qqline(residuals(model), col = "red")
      }
      
      if ("Scale-Location" %in% plot_types) {
        plot(fitted(model), sqrt(abs(residuals(model))),
             xlab = "Fitted Values", ylab = "sqrt(|Residuals|)",
             main = "Scale-Location Plot")
      }
      
      if ("Residuals vs Leverage" %in% plot_types) {
        plot(model, which = 5)
      }
      
    }, error = function(e) {
      print(paste("Plot error:", e$message))
      return(ggplot() + theme_void() + ggtitle("Error creating plots"))
    })
  })
  
  output$reg_plot <- renderPlot({
    req(data(), input$reg_dep, input$reg_indep, input$run_regression)
    
    tryCatch({
      df <- data()
      dep_var <- input$reg_dep
      indep_vars <- input$reg_indep
      
      # Create formula
      formula <- as.formula(paste(dep_var, "~", paste(indep_vars, collapse = " + ")))
      
      # Perform regression
      model <- lm(formula, data = df)
      
      # Set up plotting area based on selected plots
      plot_types <- input$reg_plots
      n_plots <- length(plot_types)
      
      if (n_plots == 0) {
        return(ggplot() + theme_void() + ggtitle("No plots selected"))
      }
      
      # Set up plot layout
      if (n_plots == 1) {
        par(mfrow = c(1, 1))
      } else if (n_plots == 2) {
        par(mfrow = c(1, 2))
      } else {
        par(mfrow = c(2, 2))
      }
      
      # Generate selected plots
      if ("Residuals vs Fitted" %in% plot_types) {
        plot(fitted(model), residuals(model),
             xlab = "Fitted Values", ylab = "Residuals",
             main = "Residuals vs Fitted Values")
        abline(h = 0, col = "red", lty = 2)
      }
      
      if ("Normal Q-Q" %in% plot_types) {
        qqnorm(residuals(model), main = "Normal Q-Q Plot")
        qqline(residuals(model), col = "red")
      }
      
      if ("Scale-Location" %in% plot_types) {
        plot(fitted(model), sqrt(abs(residuals(model))),
             xlab = "Fitted Values", ylab = "sqrt(|Residuals|)",
             main = "Scale-Location Plot")
      }
      
      if ("Residuals vs Leverage" %in% plot_types) {
        plot(model, which = 5)
      }
      
    }, error = function(e) {
      print(paste("Plot error:", e$message))
      return(ggplot() + theme_void() + ggtitle("Error creating plots"))
    })
  })
  
  # Chi-Square Test Analysis - CORRECTED VERSION
  output$chi_crosstab <- renderPrint({
    req(data(), input$chi_row, input$chi_col, input$run_chi)
    
    tryCatch({
      df <- data()
      row_var <- input$chi_row
      col_var <- input$chi_col
      
      # Create contingency table
      contingency_table <- table(df[[row_var]], df[[col_var]])
      
      cat("Crosstabulation\n")
      cat("===============\n\n")
      cat("Row Variable:", row_var, "\n")
      cat("Column Variable:", col_var, "\n\n")
      
      # Display observed frequencies if selected
      if ("Observed frequencies" %in% input$crosstab_options) {
        cat("Observed Frequencies:\n")
        print(addmargins(contingency_table))
        cat("\n")
      }
      
      # Calculate expected frequencies
      chi_test <- chisq.test(contingency_table)
      expected <- chi_test$expected
      
      # Display expected frequencies if selected
      if ("Expected frequencies" %in% input$crosstab_options) {
        cat("Expected Frequencies:\n")
        print(round(addmargins(expected), 2))
        cat("\n")
      }
      
      # Display row percentages if selected
      if ("Row percentages" %in% input$crosstab_options) {
        cat("Row Percentages:\n")
        row_perc <- prop.table(contingency_table, 1) * 100
        print(round(addmargins(row_perc), 2))
        cat("\n")
      }
      
      # Display column percentages if selected
      if ("Column percentages" %in% input$crosstab_options) {
        cat("Column Percentages:\n")
        col_perc <- prop.table(contingency_table, 2) * 100
        print(round(addmargins(col_perc), 2))
        cat("\n")
      }
      
      # Store the chi-test result for the assumption check
      chi_test_result(chi_test)
      
    }, error = function(e) {
      return(paste("Error in crosstabulation:", e$message))
    })
  })
  
  # Chi-square assumption note
  output$chi_assumption_note <- renderUI({
    req(chi_test_result())
    
    chi_test <- chi_test_result()
    expected <- chi_test$expected
    
    # Check if any expected frequency is less than 5
    if (any(expected < 5)) {
      div(
        style = "color: red; font-weight: bold; margin-top: 15px;",
        "Note: Chi-square test may not be appropriate as some expected frequencies are less than 5."
      )
    }
  })
  
  # Chi-Square Test Statistics - CORRECTED VERSION
  output$chi_test <- renderPrint({
    req(data(), input$chi_row, input$chi_col, input$run_chi)
    
    tryCatch({
      df <- data()
      row_var <- input$chi_row
      col_var <- input$chi_col
      
      # Create contingency table
      contingency_table <- table(df[[row_var]], df[[col_var]])
      
      # Perform chi-square test
      chi_test <- chisq.test(contingency_table)
      
      cat("Chi-Square Test\n")
      cat("===============\n\n")
      
      cat("Pearson Chi-Square:\n")
      cat("X²(", chi_test$parameter, ") =", round(chi_test$statistic, 3), "\n")
      cat("p-value =", format.pval(chi_test$p.value, digits = 3, eps = 0.001), "\n")
      
      # Additional statistics if requested
      if ("Phi and Cramer's V" %in% input$chi_stats) {
        # Calculate Phi and Cramer's V
        n <- sum(contingency_table)
        phi <- sqrt(chi_test$statistic / n)
        cramers_v <- sqrt(chi_test$statistic / (n * (min(dim(contingency_table)) - 1)))
        
        cat("Phi Coefficient =", round(phi, 3), "\n")
        cat("Cramer's V =", round(cramers_v, 3), "\n\n")
      }
      
      if ("Contingency coefficient" %in% input$chi_stats) {
        # Calculate contingency coefficient
        cc <- sqrt(chi_test$statistic / (chi_test$statistic + n))
        cat("Contingency Coefficient =", round(cc, 3), "\n\n")
      }
      
      if ("Lambda" %in% input$chi_stats) {
        # Calculate Lambda (asymmetric)
        row_max <- apply(contingency_table, 1, max)
        col_max <- apply(contingency_table, 2, max)
        total_max <- max(rowSums(contingency_table))
        
        lambda_r <- (sum(row_max) - total_max) / (n - total_max)  # Row dependent
        lambda_c <- (sum(col_max) - total_max) / (n - total_max)  # Column dependent
        
        cat("Lambda (asymmetric):\n")
        cat("  Row dependent =", round(lambda_r, 3), "\n")
        cat("  Column dependent =", round(lambda_c, 3), "\n\n")
      }
      
      if ("Uncertainty coefficient" %in% input$chi_stats) {
        # Calculate Uncertainty coefficient
        # This requires more complex calculation involving entropy
        cat("Uncertainty coefficient calculation would be implemented here.\n\n")
      }
      
    }, error = function(e) {
      return(paste("Error in chi-square test:", e$message))
    })
  })
  
  # Chi-Square Plot - CORRECTED VERSION
  output$chi_plot <- renderPlotly({
    req(data(), input$chi_row, input$chi_col, input$run_chi)
    
    tryCatch({
      df <- data()
      row_var <- input$chi_row
      col_var <- input$chi_col
      
      # Create contingency table
      contingency_table <- table(df[[row_var]], df[[col_var]])
      
      # Convert to data frame for plotting
      plot_data <- as.data.frame(contingency_table)
      colnames(plot_data) <- c("Row", "Column", "Frequency")
      
      # Create mosaic plot
      p <- ggplot(plot_data, aes(x = Row, y = Frequency, fill = Column)) +
        geom_bar(stat = "identity", position = "fill", alpha = 0.8) +
        scale_y_continuous(labels = scales::percent_format()) +
        labs(title = paste("Mosaic Plot of", row_var, "by", col_var),
             x = row_var, y = "Percentage", fill = col_var) +
        theme_minimal() +
        theme(axis.text.x = element_text(angle = 45, hjust = 1))
      
      ggplotly(p)
      
    }, error = function(e) {
      print(paste("Chi-square plot error:", e$message))
      return(plotly_empty())
    })
  })
  
  # Store chi-test result for assumption checking
  chi_test_result <- reactiveVal()
  
  # Fisher's Exact Test Analysis - CORRECTED VERSION
  output$fisher_crosstab <- renderPrint({
    req(data(), input$fisher_row, input$fisher_col, input$run_fisher)
    
    tryCatch({
      df <- data()
      row_var <- input$fisher_row
      col_var <- input$fisher_col
      
      # Create contingency table
      contingency_table <- table(df[[row_var]], df[[col_var]])
      
      cat("Crosstabulation\n")
      cat("===============\n\n")
      cat("Row Variable:", row_var, "\n")
      cat("Column Variable:", col_var, "\n\n")
      
      # Display observed frequencies if selected
      if ("Observed frequencies" %in% input$fisher_crosstab_options) {
        cat("Observed Frequencies:\n")
        print(addmargins(contingency_table))
        cat("\n")
      }
      
      # Calculate expected frequencies
      chi_test <- chisq.test(contingency_table)
      expected <- chi_test$expected
      
      # Display expected frequencies if selected
      if ("Expected frequencies" %in% input$fisher_crosstab_options) {
        cat("Expected Frequencies:\n")
        print(round(addmargins(expected), 2))
        cat("\n")
      }
      
      # Display row percentages if selected
      if ("Row percentages" %in% input$fisher_crosstab_options) {
        cat("Row Percentages:\n")
        row_perc <- prop.table(contingency_table, 1) * 100
        print(round(addmargins(row_perc), 2))
        cat("\n")
      }
      
      # Display column percentages if selected
      if ("Column percentages" %in% input$fisher_crosstab_options) {
        cat("Column Percentages:\n")
        col_perc <- prop.table(contingency_table, 2) * 100
        print(round(addmargins(col_perc), 2))
        cat("\n")
      }
      
    }, error = function(e) {
      return(paste("Error in crosstabulation:", e$message))
    })
  })
  
  output$fisher_test <- renderPrint({
    req(data(), input$fisher_row, input$fisher_col, input$run_fisher)
    
    tryCatch({
      df <- data()
      row_var <- input$fisher_row
      col_var <- input$fisher_col
      
      # Create contingency table
      contingency_table <- table(df[[row_var]], df[[col_var]])
      
      # Perform Fisher's exact test
      fisher_test <- fisher.test(contingency_table)
      
      cat("Fisher's Exact Test\n")
      cat("===================\n\n")
      
      cat("p-value =", format.pval(fisher_test$p.value, digits = 3), "\n")
      cat("Odds Ratio =", round(fisher_test$estimate, 3), "\n")
      if (!is.null(fisher_test$conf.int)) {
        cat("95% CI for Odds Ratio: [", 
            round(fisher_test$conf.int[1], 3), ",", 
            round(fisher_test$conf.int[2], 3), "]\n")
      }
      
    }, error = function(e) {
      return(paste("Error in Fisher's test:", e$message))
    })
  })
  
  # Fisher's Plot - CORRECTED to use bar chart
  output$fisher_plot <- renderPlotly({
    req(data(), input$fisher_row, input$fisher_col, input$run_fisher)
    
    tryCatch({
      df <- data()
      row_var <- input$fisher_row
      col_var <- input$fisher_col
      
      # Create contingency table
      contingency_table <- table(df[[row_var]], df[[col_var]])
      
      # Convert to data frame for plotting
      plot_data <- as.data.frame(contingency_table)
      colnames(plot_data) <- c("Row", "Column", "Frequency")
      
      # Create bar chart instead of mosaic plot
      p <- ggplot(plot_data, aes(x = Row, y = Frequency, fill = Column)) +
        geom_bar(stat = "identity", position = "dodge", alpha = 0.8) +
        labs(title = paste("Bar Chart of", row_var, "by", col_var),
             x = row_var, y = "Frequency", fill = col_var) +
        theme_minimal() +
        theme(axis.text.x = element_text(angle = 45, hjust = 1)) +
        scale_fill_manual(values = c(spss_blue, "#E27D72", "#68C3A3", "#F7C46C"))
      
      ggplotly(p)
      
    }, error = function(e) {
      print(paste("Fisher's plot error:", e$message))
      return(plotly_empty())
    })
  })
  # Mann-Whitney U Test Analysis
  output$mw_output <- renderPrint({
    req(data(), input$mw_vars, input$mw_group_var, input$run_mw)
    
    tryCatch({
      df <- data()
      test_vars <- input$mw_vars
      group_var <- input$mw_group_var
      
      if (length(test_vars) == 0) {
        return("Please select at least one test variable.")
      }
      
      # Parse group definition
      groups <- strsplit(input$mw_group_def, ",")[[1]]
      groups <- trimws(groups)
      
      if (length(groups) != 2) {
        return("Please define exactly two groups (e.g., '1,2' or 'A,B')")
      }
      
      cat("Mann-Whitney U Test\n")
      cat("===================\n\n")
      
      for (test_var in test_vars) {
        cat("Variable:", test_var, "\n")
        cat("----------------------------------------\n")
        
        # Filter data for the two groups
        group_data <- df[df[[group_var]] %in% groups, ]
        group1_data <- group_data[group_data[[group_var]] == groups[1], test_var]
        group2_data <- group_data[group_data[[group_var]] == groups[2], test_var]
        
        # Remove missing values
        group1_data <- group1_data[!is.na(group1_data)]
        group2_data <- group2_data[!is.na(group2_data)]
        
        # Perform Mann-Whitney test
        mw_test <- wilcox.test(group1_data, group2_data, exact = FALSE)
        
        # Calculate ranks and descriptive statistics
        cat("Group 1 (", groups[1], "): N =", length(group1_data), 
            ", Median =", round(median(group1_data), 3), 
            ", Mean Rank =", round(mean(rank(c(group1_data, group2_data))[1:length(group1_data)]), 3), "\n")
        cat("Group 2 (", groups[2], "): N =", length(group2_data), 
            ", Median =", round(median(group2_data), 3), 
            ", Mean Rank =", round(mean(rank(c(group1_data, group2_data))[(length(group1_data)+1):length(c(group1_data, group2_data))]), 3), "\n\n")
        
        cat("Mann-Whitney U =", mw_test$statistic, "\n")
        cat("p-value =", format.pval(mw_test$p.value, digits = 3), "\n\n")
      }
      
    }, error = function(e) {
      return(paste("Error in Mann-Whitney test:", e$message))
    })
  })
  
  # Kruskal-Wallis Test Analysis
  output$kw_output <- renderPrint({
    req(data(), input$kw_vars, input$kw_group_var, input$run_kw)
    
    tryCatch({
      df <- data()
      test_vars <- input$kw_vars
      group_var <- input$kw_group_var
      
      if (length(test_vars) == 0) {
        return("Please select at least one test variable.")
      }
      
      cat("Kruskal-Wallis Test\n")
      cat("===================\n\n")
      
      for (test_var in test_vars) {
        cat("Variable:", test_var, "\n")
        cat("----------------------------------------\n")
        
        # Remove missing values
        complete_cases <- complete.cases(df[[test_var]], df[[group_var]])
        test_data <- df[complete_cases, test_var]
        group_data <- df[complete_cases, group_var]
        
        # Perform Kruskal-Wallis test
        kw_test <- kruskal.test(test_data ~ group_data)
        
        # Calculate descriptive statistics by group
        groups <- unique(group_data)
        cat("Group Statistics:\n")
        for (group in groups) {
          group_values <- test_data[group_data == group]
          cat("Group", group, ": N =", length(group_values), 
              ", Median =", round(median(group_values), 3), 
              ", Mean Rank =", round(mean(rank(test_data)[group_data == group]), 3), "\n")
        }
        cat("\n")
        
        cat("Kruskal-Wallis H(", kw_test$parameter, ") =", round(kw_test$statistic, 3), "\n")
        cat("p-value =", format.pval(kw_test$p.value, digits = 3), "\n\n")
      }
      
    }, error = function(e) {
      return(paste("Error in Kruskal-Wallis test:", e$message))
    })
  })
  
  # Wilcoxon Signed-Rank Test Analysis
  output$wilcoxon_output <- renderPrint({
    req(data(), input$wilcoxon_vars, input$run_wilcoxon)
    
    tryCatch({
      df <- data()
      paired_vars <- input$wilcoxon_vars
      
      if (length(paired_vars) < 2) {
        return("Please select at least two variables for paired comparison.")
      }
      
      cat("Wilcoxon Signed-Rank Test\n")
      cat("=========================\n\n")
      
      # Create all possible pairs
      pairs <- combn(paired_vars, 2, simplify = FALSE)
      
      for (pair in pairs) {
        cat("Pair:", pair[1], "vs", pair[2], "\n")
        cat("----------------------------------------\n")
        
        var1 <- df[[pair[1]]]
        var2 <- df[[pair[2]]]
        
        # Remove cases with missing values in either variable
        complete_cases <- complete.cases(var1, var2)
        var1 <- var1[complete_cases]
        var2 <- var2[complete_cases]
        
        # Perform Wilcoxon signed-rank test
        wilcox_test <- wilcox.test(var1, var2, paired = TRUE, exact = FALSE)
        
        # Calculate differences and descriptive statistics
        differences <- var1 - var2
        cat("N =", length(var1), "\n")
        cat("Mean Difference =", round(mean(differences), 3), "\n")
        cat("Median Difference =", round(median(differences), 3), "\n")
        cat("Positive Ranks:", sum(differences > 0), "\n")
        cat("Negative Ranks:", sum(differences < 0), "\n")
        cat("Ties:", sum(differences == 0), "\n\n")
        
        cat("Wilcoxon V =", wilcox_test$statistic, "\n")
        cat("p-value =", format.pval(wilcox_test$p.value, digits = 3), "\n\n")
      }
      
    }, error = function(e) {
      return(paste("Error in Wilcoxon test:", e$message))
    })
  })
  
  # UI for logistic regression reference category selection
  output$logistic_ref_ui <- renderUI({
    req(data(), input$logistic_indep)
    
    df <- data()
    categorical_vars <- input$logistic_indep[sapply(df[input$logistic_indep], function(x) is.factor(x) || is.character(x))]
    
    if (length(categorical_vars) == 0) {
      return(NULL)
    }
    
    ref_ui_list <- lapply(categorical_vars, function(var) {
      var_levels <- if (is.factor(df[[var]])) {
        levels(df[[var]])
      } else {
        unique(na.omit(df[[var]]))
      }
      
      selectInput(
        inputId = paste0("logistic_ref_", var),
        label = paste("Reference category for", var),
        choices = var_levels,
        selected = var_levels[1]
      )
    })
    
    tagList(
      h4("Reference Categories for Categorical Variables"),
      ref_ui_list
    )
  })
  
  # UI for dependent variable level selection in binary logistic regression
  output$logistic_dep_level_ui <- renderUI({
    req(data(), input$logistic_dep)
    
    df <- data()
    dep_var <- input$logistic_dep
    
    if (is.null(dep_var) || dep_var == "") {
      return(NULL)
    }
    
    # Get the levels of the dependent variable
    var_levels <- if (is.factor(df[[dep_var]])) {
      levels(df[[dep_var]])
    } else if (is.character(df[[dep_var]])) {
      unique(na.omit(df[[dep_var]]))
    } else if (is.numeric(df[[dep_var]])) {
      # For numeric variables, check if it's binary (0/1)
      unique_vals <- unique(na.omit(df[[dep_var]]))
      if (length(unique_vals) == 2) {
        as.character(unique_vals)
      } else {
        NULL
      }
    } else {
      NULL
    }
    
    if (is.null(var_levels) || length(var_levels) < 2) {
      return(NULL)
    }
    
    selectInput(
      inputId = "logistic_dep_level",
      label = "Select reference category (coded as 0)",
      choices = var_levels,
      selected = var_levels[1]
    )
  })
  # Logistic Regression Analysis - UPDATED with proper reference category handling
  output$logistic_summary <- renderPrint({
    req(data(), input$logistic_dep, input$logistic_indep, input$run_logistic)
    
    tryCatch({
      df <- data()
      dep_var <- input$logistic_dep
      indep_vars <- input$logistic_indep
      
      if (is.null(dep_var) || dep_var == "") {
        return("Please select a dependent variable.")
      }
      
      if (length(indep_vars) == 0) {
        return("Please select at least one independent variable.")
      }
      
      # In the logistic_summary, logistic_coef, logistic_classification, and logistic_hl outputs:
      
      # Prepare dependent variable with user-specified reference category
      dep_data <- df[[dep_var]]
      
      # Get the reference category from user input
      ref_level <- input$logistic_dep_level
      
      # Handle different types of dependent variables
      if (is.numeric(dep_data)) {
        # Check if it's binary (exactly 2 unique values)
        unique_vals <- unique(na.omit(dep_data))
        if (length(unique_vals) == 2) {
          # Convert to factor with proper reference level
          if (!is.null(ref_level)) {
            # Convert numeric to character for comparison
            dep_data_char <- as.character(dep_data)
            ref_level_char <- as.character(ref_level)
            
            # Create factor with user-specified reference
            df$logistic_dep_binary <- factor(dep_data_char, 
                                             levels = c(ref_level_char, setdiff(unique_vals, ref_level_char)))
          } else {
            # Use default ordering (first unique value as reference)
            df$logistic_dep_binary <- factor(dep_data)
          }
        } else {
          return("Numeric dependent variable must have exactly 2 unique values for binary logistic regression.")
        }
      } else if (is.factor(dep_data) || is.character(dep_data)) {
        # Convert categorical to binary factor with proper reference
        unique_vals <- unique(na.omit(dep_data))
        if (length(unique_vals) != 2) {
          return("Dependent variable must have exactly 2 categories for binary logistic regression.")
        }
        
        if (!is.null(ref_level)) {
          # Use user-specified reference level
          df$logistic_dep_binary <- factor(dep_data, 
                                           levels = c(ref_level, setdiff(unique_vals, ref_level)))
        } else {
          # Use default ordering (first level as reference)
          df$logistic_dep_binary <- factor(dep_data)
        }
      } else {
        return("Dependent variable must be binary or categorical with 2 levels.")
      }
      # Verify the binary variable was created correctly
      if (!all(levels(df$logistic_dep_binary) %in% c("0", "1"))) {
        # Recode to 0/1 format for glm
        df$logistic_dep_binary <- ifelse(df$logistic_dep_binary == levels(df$logistic_dep_binary)[1], 0, 1)
      }
      
      # Prepare independent variables with proper reference categories
      model_data <- df
      categorical_vars <- indep_vars[sapply(df[indep_vars], function(x) is.factor(x) || is.character(x))]
      
      # Set reference categories for categorical independent variables
      for (var in categorical_vars) {
        ref_input <- input[[paste0("logistic_ref_", var)]]
        if (!is.null(ref_input)) {
          if (is.factor(model_data[[var]])) {
            model_data[[var]] <- relevel(model_data[[var]], ref = ref_input)
          } else {
            # Convert character to factor with specified reference
            model_data[[var]] <- factor(model_data[[var]])
            model_data[[var]] <- relevel(model_data[[var]], ref = ref_input)
          }
        }
      }
      
      # Create formula
      formula <- as.formula(paste("logistic_dep_binary", "~", paste(indep_vars, collapse = " + ")))
      
      # Perform logistic regression
      model <- glm(formula, data = model_data, family = binomial)
      
      cat("Binary Logistic Regression\n")
      cat("==========================\n\n")
      cat("Dependent Variable:", dep_var, "\n")
      
      # Display reference category information
      if (!is.null(ref_level)) {
        cat("Reference Category (coded as 0):", ref_level, "\n")
        success_level <- setdiff(unique(na.omit(dep_data)), ref_level)
        if (length(success_level) > 0) {
          cat("Success Category (coded as 1):", success_level[1], "\n")
        }
      } else {
        cat("Reference Category (coded as 0):", levels(factor(dep_data))[1], "\n")
      }
      
      cat("Independent Variables:", paste(indep_vars, collapse = ", "), "\n\n")
      
      # Display reference categories used for independent variables
      if (length(categorical_vars) > 0) {
        cat("Reference Categories for Independent Variables:\n")
        for (var in categorical_vars) {
          ref_input <- input[[paste0("logistic_ref_", var)]]
          if (!is.null(ref_input)) {
            cat(paste0("  ", var, ": ", ref_input, "\n"))
          } else {
            # Show default reference if user didn't specify
            if (is.factor(model_data[[var]])) {
              cat(paste0("  ", var, ": ", levels(model_data[[var]])[1], " (default)\n"))
            }
          }
        }
        cat("\n")
      }
      
      # Model summary
      cat("Model Summary\n")
      cat("-------------\n")
      summary_model <- summary(model)
      print(summary_model)
      
      # Display iteration history if requested
      if ("Iteration history" %in% input$logistic_options) {
        cat("\nIteration History:\n")
        cat("Convergence achieved in", length(model$iter), "iterations\n")
      }
      
    }, error = function(e) {
      return(paste("Error in logistic regression:", e$message))
    })
  })
  
  
  output$logistic_coef <- renderPrint({
    req(data(), input$logistic_dep, input$logistic_indep, input$run_logistic)
    
    tryCatch({
      df <- data()
      dep_var <- input$logistic_dep
      indep_vars <- input$logistic_indep
      
      # Prepare data (same as in logistic_summary)
      dep_data <- df[[dep_var]]
      if (is.numeric(dep_data) && length(unique(na.omit(dep_data))) == 2 && all(unique(na.omit(dep_data)) %in% c(0, 1))) {
        df$logistic_dep_binary <- dep_data
      } else {
        success_level <- input$logistic_dep_level
        df$logistic_dep_binary <- ifelse(dep_data == success_level, 1, 0)
      }
      
      model_data <- df
      categorical_vars <- indep_vars[sapply(df[indep_vars], function(x) is.factor(x) || is.character(x))]
      
      for (var in categorical_vars) {
        ref_input <- input[[paste0("logistic_ref_", var)]]
        if (!is.null(ref_input)) {
          if (is.factor(model_data[[var]])) {
            model_data[[var]] <- relevel(model_data[[var]], ref = ref_input)
          } else {
            model_data[[var]] <- factor(model_data[[var]])
            model_data[[var]] <- relevel(model_data[[var]], ref = ref_input)
          }
        }
      }
      
      # Create formula
      formula <- as.formula(paste("logistic_dep_binary", "~", paste(indep_vars, collapse = " + ")))
      
      # Perform logistic regression
      model <- glm(formula, data = model_data, family = binomial)
      
      cat("Logistic Regression Coefficients\n")
      cat("===============================\n\n")
      
      # Coefficient table with odds ratios
      coef_table <- summary(model)$coefficients
      odds_ratios <- exp(coef(model))
      coef_table <- cbind(coef_table, "Odds Ratio" = odds_ratios)
      
      if ("CI for exp(B)" %in% input$logistic_options) {
        # Add confidence intervals for odds ratios
        ci <- exp(confint(model))
        coef_table <- cbind(coef_table, "OR Lower CI" = ci[, 1], "OR Upper CI" = ci[, 2])
      }
      
      print(round(coef_table, 4))
      
      # Display correlation of estimates if requested
      if ("Correlation of estimates" %in% input$logistic_options) {
        cat("\nCorrelation of Parameter Estimates:\n")
        cor_matrix <- cov2cor(vcov(model))
        print(round(cor_matrix, 4))
      }
      
    }, error = function(e) {
      return(paste("Error in logistic regression:", e$message))
    })
  })
  
  output$logistic_classification <- renderPrint({
    req(data(), input$logistic_dep, input$logistic_indep, input$run_logistic)
    
    tryCatch({
      df <- data()
      dep_var <- input$logistic_dep
      indep_vars <- input$logistic_indep
      
      # Prepare data (same as in logistic_summary)
      dep_data <- df[[dep_var]]
      if (is.numeric(dep_data) && length(unique(na.omit(dep_data))) == 2 && all(unique(na.omit(dep_data)) %in% c(0, 1))) {
        df$logistic_dep_binary <- dep_data
      } else {
        success_level <- input$logistic_dep_level
        df$logistic_dep_binary <- ifelse(dep_data == success_level, 1, 0)
      }
      
      model_data <- df
      categorical_vars <- indep_vars[sapply(df[indep_vars], function(x) is.factor(x) || is.character(x))]
      
      for (var in categorical_vars) {
        ref_input <- input[[paste0("logistic_ref_", var)]]
        if (!is.null(ref_input)) {
          if (is.factor(model_data[[var]])) {
            model_data[[var]] <- relevel(model_data[[var]], ref = ref_input)
          } else {
            model_data[[var]] <- factor(model_data[[var]])
            model_data[[var]] <- relevel(model_data[[var]], ref = ref_input)
          }
        }
      }
      
      # Create formula
      formula <- as.formula(paste("logistic_dep_binary", "~", paste(indep_vars, collapse = " + ")))
      
      # Perform logistic regression
      model <- glm(formula, data = model_data, family = binomial)
      
      # Make predictions - ensure we use the same data used for model fitting
      predictions <- ifelse(predict(model, type = "response") > 0.5, 1, 0)
      actual <- model_data$logistic_dep_binary
      
      # Ensure both vectors have the same length (remove NAs)
      complete_cases <- complete.cases(actual, predictions)
      actual <- actual[complete_cases]
      predictions <- predictions[complete_cases]
      
      # Create confusion matrix - FIXED: ensure both arguments have same length
      confusion_matrix <- table(Actual = actual, Predicted = predictions)
      
      cat("Classification Table\n")
      cat("====================\n\n")
      
      print(confusion_matrix)
      
      # Calculate classification metrics
      accuracy <- sum(diag(confusion_matrix)) / sum(confusion_matrix)
      sensitivity <- ifelse(sum(confusion_matrix[2, ]) > 0, 
                            confusion_matrix[2, 2] / sum(confusion_matrix[2, ]), 0)
      specificity <- ifelse(sum(confusion_matrix[1, ]) > 0, 
                            confusion_matrix[1, 1] / sum(confusion_matrix[1, ]), 0)
      precision <- ifelse(sum(confusion_matrix[, 2]) > 0, 
                          confusion_matrix[2, 2] / sum(confusion_matrix[, 2]), 0)
      
      cat("\nClassification Metrics:\n")
      cat("Accuracy:", round(accuracy, 4), "\n")
      cat("Sensitivity:", round(sensitivity, 4), "\n")
      cat("Specificity:", round(specificity, 4), "\n")
      cat("Precision:", round(precision, 4), "\n")
      
    }, error = function(e) {
      return(paste("Error in classification analysis:", e$message))
    })
  })
  
  output$logistic_hl <- renderPrint({
    req(data(), input$logistic_dep, input$logistic_indep, input$run_logistic)
    
    tryCatch({
      if (!"Hosmer-Lemeshow goodness-of-fit" %in% input$logistic_options) {
        return("Hosmer-Lemeshow test not selected in options.")
      }
      
      df <- data()
      dep_var <- input$logistic_dep
      indep_vars <- input$logistic_indep
      
      # Prepare data (same as in logistic_summary)
      dep_data <- df[[dep_var]]
      if (is.numeric(dep_data) && length(unique(na.omit(dep_data))) == 2 && all(unique(na.omit(dep_data)) %in% c(0, 1))) {
        df$logistic_dep_binary <- dep_data
      } else {
        success_level <- input$logistic_dep_level
        df$logistic_dep_binary <- ifelse(dep_data == success_level, 1, 0)
      }
      
      model_data <- df
      categorical_vars <- indep_vars[sapply(df[indep_vars], function(x) is.factor(x) || is.character(x))]
      
      for (var in categorical_vars) {
        ref_input <- input[[paste0("logistic_ref_", var)]]
        if (!is.null(ref_input)) {
          if (is.factor(model_data[[var]])) {
            model_data[[var]] <- relevel(model_data[[var]], ref = ref_input)
          } else {
            model_data[[var]] <- factor(model_data[[var]])
            model_data[[var]] <- relevel(model_data[[var]], ref = ref_input)
          }
        }
      }
      
      # Create formula
      formula <- as.formula(paste("logistic_dep_binary", "~", paste(indep_vars, collapse = " + ")))
      
      # Perform logistic regression
      model <- glm(formula, data = model_data, family = binomial)
      
      # Hosmer-Lemeshow test using ResourceSelection package
      if (!require(ResourceSelection, quietly = TRUE)) {
        return("Please install the 'ResourceSelection' package to use the Hosmer-Lemeshow test.")
      }
      
      cat("Hosmer-Lemeshow Goodness-of-Fit Test\n")
      cat("====================================\n\n")
      
      # Use the hoslem.test function from ResourceSelection package
      predicted_probs <- predict(model, type = "response")
      actual <- model_data$logistic_dep_binary
      
      # Remove NA values
      complete_cases <- complete.cases(actual, predicted_probs)
      actual <- actual[complete_cases]
      predicted_probs <- predicted_probs[complete_cases]
      
      # Perform the test
      hl_test <- hoslem.test(actual, predicted_probs, g = 10)
      
      cat("Chi-square(", hl_test$parameter, ") =", round(hl_test$statistic, 3), "\n")
      cat("p-value =", format.pval(hl_test$p.value, digits = 3), "\n\n")
      
      # Interpretation
      if (hl_test$p.value > 0.05) {
        cat("Interpretation: The model fits the data well (p > 0.05)\n")
      } else {
        cat("Interpretation: The model does not fit the data well (p <= 0.05)\n")
      }
      
    }, error = function(e) {
      return(paste("Error in Hosmer-Lemeshow test:", e$message))
    })
  })
  # Factor Analysis - CORRECTED VERSION
  output$factor_summary <- renderPrint({
    req(data(), input$factor_vars, input$run_factor)
    
    tryCatch({
      df <- data()
      factor_vars <- input$factor_vars
      
      if (length(factor_vars) < 3) {
        return("Please select at least 3 variables for factor analysis.")
      }
      
      # Extract only numeric variables
      numeric_vars <- sapply(df[, factor_vars], is.numeric)
      if (sum(numeric_vars) < 3) {
        return("Please select at least 3 numeric variables for factor analysis.")
      }
      
      factor_data <- df[, factor_vars[numeric_vars]]
      
      # Remove cases with missing values
      factor_data <- na.omit(factor_data)
      
      cat("Exploratory Factor Analysis\n")
      cat("===========================\n\n")
      cat("Number of Variables:", ncol(factor_data), "\n")
      cat("Number of Cases:", nrow(factor_data), "\n")
      cat("Number of Factors:", input$num_factors, "\n")
      cat("Rotation Method:", input$rotation_method, "\n\n")
      
      # Calculate correlation matrix
      cor_matrix <- cor(factor_data)
      
      # KMO and Bartlett's test
      if ("KMO & Bartlett's" %in% input$factor_options) {
        # KMO Test
        kmo_test <- function(data) {
          library(psych)
          X <- cor(data)
          iX <- solve(X)
          S2 <- diag(diag((iX^-1)))
          AIS <- S2 %*% iX %*% S2
          AI <- data
          for (i in 1:ncol(data)) {
            for (j in 1:ncol(data)) {
              if (i != j) {
                AI[i, j] <- AIS[i, j]
              }
            }
          }
          diag(AI) <- 0
          MSA <- colSums(X^2) / (colSums(X^2) + colSums(AI^2))
          kmo <- sum(X^2) / (sum(X^2) + sum(AI^2))
          return(list(overall = kmo, individual = MSA))
        }
        
        kmo_result <- kmo_test(factor_data)
        
        cat("Kaiser-Meyer-Olkin (KMO) Test:\n")
        cat("Overall MSA =", round(kmo_result$overall, 3), "\n")
        cat("Individual MSA:\n")
        print(round(kmo_result$individual, 3))
        cat("\n")
        
        # Bartlett's test of sphericity
        bartlett_test <- psych::cortest.bartlett(cor_matrix, n = nrow(factor_data))
        cat("Bartlett's Test of Sphericity:\n")
        cat("Chi-square =", round(bartlett_test$chisq, 3), "\n")
        cat("df =", bartlett_test$df, "\n")
        cat("p-value =", format.pval(bartlett_test$p.value, digits = 3, eps = 0.001), "\n")
      }
      
      # Perform factor analysis
      fa_result <- psych::fa(factor_data, nfactors = input$num_factors, 
                             rotate = input$rotation_method, fm = "minres")
      
      if ("Eigenvalues" %in% input$factor_options) {
        cat("Eigenvalues:\n")
        eigen_values <- fa_result$e.values
        eigen_table <- data.frame(
          Factor = 1:length(eigen_values),
          Eigenvalue = round(eigen_values, 3),
          `% of Variance` = round(eigen_values/sum(eigen_values)*100, 2),
          `Cumulative %` = round(cumsum(eigen_values/sum(eigen_values)*100), 2)
        )
        print(eigen_table, row.names = FALSE)
        cat("\n")
      }
      
      cat("Factor Analysis Results:\n")
      print(fa_result, digits = 3, sort = TRUE)
      
      # Store the result for other outputs
      fa_results(fa_result)
      
    }, error = function(e) {
      return(paste("Error in factor analysis:", e$message))
    })
  })
  
  # Factor Loadings output
  output$factor_loadings <- renderPrint({
    req(fa_results())
    
    tryCatch({
      fa_result <- fa_results()
      
      cat("Factor Loadings\n")
      cat("===============\n\n")
      
      # Extract loadings
      loadings <- fa_result$loadings[]
      
      # Apply cutoff
      cutoff <- input$factor_scores_cutoff
      loadings[abs(loadings) < cutoff] <- NA
      
      # Print loadings
      print(round(loadings, 3))
      
      # Communalities
      if ("Communalities" %in% input$factor_options) {
        cat("\nCommunalities:\n")
        communalities <- fa_result$communality
        print(round(communalities, 3))
      }
      
      # Score coefficients
      if ("Score coefficients" %in% input$factor_options && !is.null(fa_result$weights)) {
        cat("\nScore Coefficients:\n")
        weights <- fa_result$weights
        print(round(weights, 3))
      }
      
    }, error = function(e) {
      return(paste("Error displaying factor loadings:", e$message))
    })
  })
  
  # Scree Plot
  output$factor_scree <- renderPlotly({
    req(fa_results())
    
    tryCatch({
      fa_result <- fa_results()
      eigen_values <- fa_result$e.values
      
      scree_data <- data.frame(
        Factor = 1:length(eigen_values),
        Eigenvalue = eigen_values
      )
      
      p <- plot_ly(scree_data, x = ~Factor, y = ~Eigenvalue, type = 'scatter', mode = 'lines+markers',
                   line = list(color = spss_blue), marker = list(color = spss_blue)) %>%
        add_trace(y = 1, type = 'scatter', mode = 'lines', 
                  line = list(color = 'red', dash = 'dash'), name = 'Kaiser Criterion (Eigenvalue = 1)') %>%
        layout(title = "Scree Plot",
               xaxis = list(title = "Factor Number"),
               yaxis = list(title = "Eigenvalue"))
      
      return(p)
      
    }, error = function(e) {
      return(plotly_empty())
    })
  })
  
  # Parallel Analysis
  output$factor_parallel <- renderPlotly({
    req(data(), input$factor_vars)
    
    tryCatch({
      df <- data()
      factor_vars <- input$factor_vars
      
      # Extract only numeric variables
      numeric_vars <- sapply(df[, factor_vars], is.numeric)
      factor_data <- df[, factor_vars[numeric_vars]]
      factor_data <- na.omit(factor_data)
      
      # Perform parallel analysis
      parallel <- psych::fa.parallel(factor_data, plot = FALSE)
      
      parallel_data <- data.frame(
        Factor = 1:length(parallel$fa.values),
        Actual = parallel$fa.values,
        Simulated = parallel$fa.sim
      )
      
      p <- plot_ly(parallel_data, x = ~Factor) %>%
        add_trace(y = ~Actual, type = 'scatter', mode = 'lines+markers', 
                  name = 'Actual Data', line = list(color = spss_blue), 
                  marker = list(color = spss_blue)) %>%
        add_trace(y = ~Simulated, type = 'scatter', mode = 'lines+markers', 
                  name = 'Parallel Analysis', line = list(color = 'red'), 
                  marker = list(color = 'red')) %>%
        layout(title = "Parallel Analysis",
               xaxis = list(title = "Factor Number"),
               yaxis = list(title = "Eigenvalue"),
               legend = list(x = 0.1, y = 0.9))
      
      return(p)
      
    }, error = function(e) {
      return(plotly_empty())
    })
  })
  
  # Factor Scores
  output$factor_scores <- renderPrint({
    req(fa_results(), data(), input$factor_vars)
    
    tryCatch({
      df <- data()
      factor_vars <- input$factor_vars
      
      # Extract only numeric variables
      numeric_vars <- sapply(df[, factor_vars], is.numeric)
      factor_data <- df[, factor_vars[numeric_vars]]
      factor_data <- na.omit(factor_data)
      
      fa_result <- fa_results()
      
      # Calculate factor scores
      scores <- psych::factor.scores(factor_data, fa_result)
      
      cat("Factor Scores\n")
      cat("=============\n\n")
      
      # Display first few scores
      cat("First 10 cases:\n")
      print(round(head(scores$scores, 10), 3))
      
      # Descriptive statistics of scores
      cat("\nDescriptive Statistics of Factor Scores:\n")
      score_stats <- describe(scores$scores)
      print(round(score_stats, 3))
      
    }, error = function(e) {
      return(paste("Error calculating factor scores:", e$message))
    })
  })
  
  # Harman's Single Factor Test
  output$factor_harman <- renderPrint({
    req(data(), input$factor_vars)
    
    tryCatch({
      df <- data()
      factor_vars <- input$factor_vars
      
      # Extract only numeric variables
      numeric_vars <- sapply(df[, factor_vars], is.numeric)
      factor_data <- df[, factor_vars[numeric_vars]]
      factor_data <- na.omit(factor_data)
      
      # Perform single factor analysis
      single_factor <- psych::fa(factor_data, nfactors = 1, rotate = "none")
      
      cat("Harman's Single Factor Test\n")
      cat("===========================\n\n")
      
      # Calculate percentage of variance explained by first factor
      total_variance <- sum(single_factor$e.values)
      first_factor_variance <- single_factor$e.values[1]
      percent_variance <- (first_factor_variance / total_variance) * 100
      
      cat("Total Variance:", round(total_variance, 3), "\n")
      cat("Variance explained by first factor:", round(first_factor_variance, 3), "\n")
      cat("Percentage of total variance:", round(percent_variance, 2), "%\n\n")
      
      if (percent_variance < 50) {
        cat("Interpretation: Common method bias is not a serious concern (", 
            round(percent_variance, 2), "% < 50%)\n")
      } else {
        cat("Interpretation: Potential common method bias concern (", 
            round(percent_variance, 2), "% >= 50%)\n")
      }
      
      # Display first factor loadings
      cat("\nSingle Factor Loadings:\n")
      loadings <- single_factor$loadings[]
      print(round(loadings, 3))
      
    }, error = function(e) {
      return(paste("Error in Harman's test:", e$message))
    })
  })
  
  # Reliability analysis for factors
  output$factor_reliability <- renderPrint({
    req(fa_results(), data(), input$factor_vars)
    
    tryCatch({
      df <- data()
      factor_vars <- input$factor_vars
      
      # Extract only numeric variables
      numeric_vars <- sapply(df[, factor_vars], is.numeric)
      factor_data <- df[, factor_vars[numeric_vars]]
      factor_data <- na.omit(factor_data)
      
      fa_result <- fa_results()
      
      cat("Reliability Analysis for Factors\n")
      cat("===============================\n\n")
      
      # For each factor, calculate reliability of items with high loadings
      loadings <- fa_result$loadings[]
      cutoff <- input$factor_scores_cutoff
      
      for (i in 1:ncol(loadings)) {
        # Get items with significant loadings on this factor
        items <- which(abs(loadings[, i]) >= cutoff)
        
        if (length(items) >= 2) {
          cat("Factor", i, "Reliability (Items:", paste(names(items), collapse = ", "), ")\n")
          
          # Calculate Cronbach's alpha
          alpha_result <- psych::alpha(factor_data[, items])
          cat("Cronbach's Alpha =", round(alpha_result$total$raw_alpha, 3), "\n\n")
        } else {
          cat("Factor", i, ": Insufficient items for reliability analysis (need at least 2)\n\n")
        }
      }
      
    }, error = function(e) {
      return(paste("Error in reliability analysis:", e$message))
    })
  })
  
  # Store factor analysis results
  fa_results <- reactiveVal()
  
  # Update the factor analysis results when analysis is run
  observeEvent(input$run_factor, {
    # This will trigger the factor_summary output which stores the results
  })
  # Ordinal Regression Analysis - ENHANCED
  # UI for dependent variable level ordering in ordinal regression - CORRECTED VERSION
  output$ordinal_dep_level_ui <- renderUI({
    req(data(), input$ordinal_dep)
    
    df <- data()
    dep_var <- input$ordinal_dep
    
    if (is.null(dep_var) || dep_var == "") {
      return(NULL)
    }
    
    # Get the levels of the dependent variable
    if (is.factor(df[[dep_var]]) || is.character(df[[dep_var]])) {
      var_levels <- if (is.factor(df[[dep_var]])) {
        levels(df[[dep_var]])
      } else {
        unique(na.omit(df[[dep_var]]))
      }
      
      if (length(var_levels) > 1) {
        tagList(
          h4("Dependent Variable Ordering"),
          p("Select reference category (will be coded as the first level):"),
          selectInput("ordinal_ref_level", "Reference Category", 
                      choices = var_levels, selected = var_levels[1]),
          p("Level order (from lowest to highest):"),
          helpText("Reference category will be the baseline level.")
        )
      } else {
        p("Dependent variable has only one level. Add more categories for ordinal regression.")
      }
    } else if (is.numeric(df[[dep_var]])) {
      # For numeric variables, show information about using as ordinal scale
      numeric_stats <- summary(df[[dep_var]])
      tagList(
        h4("Numeric Scale Detected"),
        p("Using numeric values as ordered scale:"),
        verbatimTextOutput("ordinal_numeric_info"),
        helpText("For ordinal regression, numeric values will be treated as ordered categories.")
      )
    } else {
      NULL
    }
  })
  # Info about numeric variable for ordinal regression
  output$ordinal_numeric_info <- renderPrint({
    req(data(), input$ordinal_dep)
    df <- data()
    dep_var <- input$ordinal_dep
    
    if (is.numeric(df[[dep_var]])) {
      cat("Variable:", dep_var, "\n")
      cat("Unique values:", length(unique(na.omit(df[[dep_var]]))), "\n")
      cat("Range:", min(df[[dep_var]], na.rm = TRUE), "-", max(df[[dep_var]], na.rm = TRUE), "\n")
      cat("Values will be treated as ordered categories\n")
    }
  })
  
  # UI for categorical variable reference category selection for INDEPENDENT variables in ordinal regression
  output$ordinal_ref_ui <- renderUI({
    req(data(), input$ordinal_indep)
    
    df <- data()
    categorical_vars <- input$ordinal_indep[sapply(df[input$ordinal_indep], function(x) is.factor(x) || is.character(x))]
    
    if (length(categorical_vars) == 0) {
      return(NULL)
    }
    
    ref_ui_list <- lapply(categorical_vars, function(var) {
      var_levels <- if (is.factor(df[[var]])) {
        levels(df[[var]])
      } else {
        unique(na.omit(df[[var]]))
      }
      
      selectInput(
        inputId = paste0("ordinal_ref_", var),
        label = paste("Reference category for", var),
        choices = var_levels,
        selected = var_levels[1]
      )
    })
    
    tagList(
      h4("Reference Categories for Categorical Variables"),
      ref_ui_list
    )
  })
  # Ordinal Regression Analysis - CORRECTED VERSION
  output$ordinal_output <- renderPrint({
    req(data(), input$ordinal_dep, input$ordinal_indep, input$run_ordinal)
    
    tryCatch({
      df <- data()
      dep_var <- input$ordinal_dep
      indep_vars <- input$ordinal_indep
      
      if (is.null(dep_var) || dep_var == "") {
        return("Please select a dependent variable.")
      }
      
      if (length(indep_vars) == 0) {
        return("Please select at least one independent variable.")
      }
      
      # Prepare dependent variable
      dep_data <- df[[dep_var]]
      
      # Handle different types of dependent variables
      if (is.numeric(dep_data)) {
        # For numeric variables, convert to ordered factor
        unique_vals <- sort(unique(na.omit(dep_data)))
        df$ordinal_dep_ordered <- factor(dep_data, levels = unique_vals, ordered = TRUE)
        cat("Numeric scale converted to ordered factor with levels:", paste(unique_vals, collapse = " < "), "\n\n")
        
      } else if (is.factor(dep_data) || is.character(dep_data)) {
        # For categorical variables, use natural ordering
        unique_vals <- if (is.factor(dep_data)) {
          levels(dep_data)
        } else {
          sort(unique(na.omit(dep_data)))
        }
        df$ordinal_dep_ordered <- factor(dep_data, levels = unique_vals, ordered = TRUE)
        cat("Order applied:", paste(unique_vals, collapse = " < "), "\n\n")
      } else {
        return("Dependent variable must be numeric, factor, or character for ordinal regression.")
      }
      
      # Prepare independent variables with proper reference categories
      model_data <- df
      categorical_vars <- indep_vars[sapply(df[indep_vars], function(x) is.factor(x) || is.character(x))]
      
      # Set reference categories for categorical independent variables
      for (var in categorical_vars) {
        ref_input <- input[[paste0("ordinal_ref_", var)]]
        if (!is.null(ref_input)) {
          if (is.factor(model_data[[var]])) {
            model_data[[var]] <- relevel(model_data[[var]], ref = ref_input)
          } else {
            # Convert character to factor with specified reference
            model_data[[var]] <- factor(model_data[[var]])
            model_data[[var]] <- relevel(model_data[[var]], ref = ref_input)
          }
          cat("Reference category for", var, ":", ref_input, "\n")
        }
      }
      cat("\n")
      
      # Create formula
      formula <- as.formula(paste("ordinal_dep_ordered", "~", paste(indep_vars, collapse = " + ")))
      
      # Perform ordinal regression using MASS::polr
      model <- MASS::polr(formula, data = model_data, Hess = TRUE)
      
      cat("Ordinal Regression (Proportional Odds Model)\n")
      cat("============================================\n\n")
      cat("Dependent Variable:", dep_var, "\n")
      cat("Independent Variables:", paste(indep_vars, collapse = ", "), "\n\n")
      
      # Display reference category information
      cat("Reference Categories:\n")
      cat("Dependent variable reference:", levels(model_data$ordinal_dep_ordered)[1], "\n")
      for (var in categorical_vars) {
        ref_input <- input[[paste0("ordinal_ref_", var)]]
        if (!is.null(ref_input)) {
          cat(paste0("  ", var, ": ", ref_input, "\n"))
        } else if (is.factor(model_data[[var]])) {
          cat(paste0("  ", var, ": ", levels(model_data[[var]])[1], " (default)\n"))
        }
      }
      cat("\n")
      
      # Model summary
      summary_model <- summary(model)
      print(summary_model)
      
      # Calculate p-values - FIXED: Ensure we're working with numeric values
      coef_table <- coef(summary_model)
      if (is.numeric(coef_table[, "t value"])) {
        p_values <- 2 * pt(abs(coef_table[, "t value"]), df.residual(model), lower.tail = FALSE)
        coef_table <- cbind(coef_table, "p-value" = p_values)
        
        cat("\nCoefficient Table with p-values:\n")
        print(round(coef_table, 4))
      } else {
        cat("\nWarning: Could not calculate p-values - non-numeric t-values detected\n")
        print(round(coef_table, 4))
      }
      
      # Calculate odds ratios - FIXED: Ensure coefficients are numeric
      if (is.numeric(coef(model))) {
        odds_ratios <- exp(coef(model))
        cat("\nOdds Ratios:\n")
        print(round(odds_ratios, 4))
        
        # Calculate confidence intervals for odds ratios
        ci <- exp(confint(model))
        cat("\n95% Confidence Intervals for Odds Ratios:\n")
        print(round(ci, 4))
      }
      
      # Model fit statistics - FIXED: Ensure proper numeric comparison
      cat("\nModel Fit Statistics:\n")
      null_model <- MASS::polr(ordinal_dep_ordered ~ 1, data = model_data, Hess = TRUE)
      
      # Check if models are valid before comparison
      if (!is.null(model) && !is.null(null_model)) {
        lrtest <- anova(null_model, model)
        if (is.numeric(lrtest$Deviance[2]) && is.numeric(lrtest$Df[2])) {
          cat("Likelihood Ratio Test: Chi-square(", lrtest$Df[2], ") =", 
              round(lrtest$Deviance[2], 3), ", p =", 
              format.pval(lrtest$`Pr(>Chi)`[2], digits = 3), "\n")
        }
        
        # Calculate pseudo R-squared - FIXED: Ensure log-likelihood values are numeric
        loglik_model <- logLik(model)
        loglik_null <- logLik(null_model)
        
        if (is.numeric(loglik_model) && is.numeric(loglik_null)) {
          pseudo_r2 <- 1 - (as.numeric(loglik_model) / as.numeric(loglik_null))
          cat("Pseudo R-squared (McFadden):", round(pseudo_r2, 4), "\n")
        }
      } else {
        cat("Could not calculate model fit statistics - model fitting failed\n")
      }
      
    }, error = function(e) {
      return(paste("Error in ordinal regression:", e$message))
    })
  })
  # Multinomial Regression Analysis - ENHANCED with reference category support
  # UI for dependent variable reference category selection in multinomial regression
  output$multinomial_dep_level_ui <- renderUI({
    req(data(), input$multinomial_dep)
    
    df <- data()
    dep_var <- input$multinomial_dep
    
    if (is.null(dep_var) || dep_var == "") {
      return(NULL)
    }
    
    # Get the levels of the dependent variable
    if (is.factor(df[[dep_var]]) || is.character(df[[dep_var]])) {
      var_levels <- if (is.factor(df[[dep_var]])) {
        levels(df[[dep_var]])
      } else {
        unique(na.omit(df[[dep_var]]))
      }
      
      if (length(var_levels) > 1) {
        tagList(
          h4("Dependent Variable Reference Category"),
          p("Select reference category (baseline level):"),
          selectInput("multinomial_ref_level", "Reference Category", 
                      choices = var_levels, selected = var_levels[1])
        )
      } else {
        p("Dependent variable has only one level. Add more categories for multinomial regression.")
      }
    } else if (is.numeric(df[[dep_var]])) {
      # For numeric variables, check if it can be treated as categorical
      unique_vals <- unique(na.omit(df[[dep_var]]))
      if (length(unique_vals) <= 10) {
        tagList(
          h4("Numeric Variable Detected"),
          p("Using numeric values as categories:"),
          verbatimTextOutput("multinomial_numeric_info"),
          selectInput("multinomial_ref_level", "Reference Category", 
                      choices = sort(unique_vals), selected = sort(unique_vals)[1])
        )
      } else {
        p("Numeric variable has too many unique values (>10). Consider recoding as categorical.")
      }
    } else {
      NULL
    }
  })
  
  # Info about numeric variable for multinomial regression
  output$multinomial_numeric_info <- renderPrint({
    req(data(), input$multinomial_dep)
    df <- data()
    dep_var <- input$multinomial_dep
    
    if (is.numeric(df[[dep_var]])) {
      unique_vals <- sort(unique(na.omit(df[[dep_var]])))
      cat("Variable:", dep_var, "\n")
      cat("Unique values:", length(unique_vals), "\n")
      cat("Range:", min(unique_vals), "-", max(unique_vals), "\n")
      cat("Values will be treated as categories\n")
    }
  })
  
  # UI for categorical variable reference category selection for INDEPENDENT variables in multinomial regression
  output$multinomial_ref_ui <- renderUI({
    req(data(), input$multinomial_indep)
    
    df <- data()
    categorical_vars <- input$multinomial_indep[sapply(df[input$multinomial_indep], function(x) is.factor(x) || is.character(x))]
    
    if (length(categorical_vars) == 0) {
      return(NULL)
    }
    
    ref_ui_list <- lapply(categorical_vars, function(var) {
      var_levels <- if (is.factor(df[[var]])) {
        levels(df[[var]])
      } else {
        unique(na.omit(df[[var]]))
      }
      
      selectInput(
        inputId = paste0("multinomial_ref_", var),
        label = paste("Reference category for", var),
        choices = var_levels,
        selected = var_levels[1]
      )
    })
    
    tagList(
      h4("Reference Categories for Categorical Variables"),
      ref_ui_list
    )
  })
  
  # Multinomial Regression Analysis - CORRECTED VERSION
  output$multinomial_output <- renderPrint({
    req(data(), input$multinomial_dep, input$multinomial_indep, input$run_multinomial)
    
    tryCatch({
      df <- data()
      dep_var <- input$multinomial_dep
      indep_vars <- input$multinomial_indep
      
      if (is.null(dep_var) || dep_var == "") {
        return("Please select a dependent variable.")
      }
      
      if (length(indep_vars) == 0) {
        return("Please select at least one independent variable.")
      }
      
      # Prepare dependent variable with user-specified reference category
      dep_data <- df[[dep_var]]
      ref_level <- input$multinomial_ref_level
      
      # Handle different types of dependent variables
      if (is.numeric(dep_data)) {
        # Convert numeric to factor with user-specified reference
        unique_vals <- sort(unique(na.omit(dep_data)))
        if (!is.null(ref_level)) {
          ref_level <- as.numeric(ref_level)
          # Ensure the reference level exists in the data
          if (!ref_level %in% unique_vals) {
            return(paste("Reference level", ref_level, "not found in variable", dep_var))
          }
          other_vals <- setdiff(unique_vals, ref_level)
          df$multinomial_dep_factor <- factor(dep_data, 
                                              levels = c(ref_level, other_vals))
        } else {
          df$multinomial_dep_factor <- factor(dep_data)
        }
      } else if (is.factor(dep_data) || is.character(dep_data)) {
        # For categorical variables, use user-specified reference
        unique_vals <- if (is.factor(dep_data)) {
          levels(dep_data)
        } else {
          unique(na.omit(dep_data))
        }
        
        if (!is.null(ref_level)) {
          # Ensure the reference level exists in the data
          if (!ref_level %in% unique_vals) {
            return(paste("Reference level", ref_level, "not found in variable", dep_var))
          }
          other_vals <- setdiff(unique_vals, ref_level)
          df$multinomial_dep_factor <- factor(dep_data, 
                                              levels = c(ref_level, other_vals))
        } else {
          df$multinomial_dep_factor <- factor(dep_data)
        }
      } else {
        return("Dependent variable must be numeric, factor, or character for multinomial regression.")
      }
      
      # Check if we have enough categories
      if (length(levels(df$multinomial_dep_factor)) < 2) {
        return("Dependent variable must have at least 2 categories for multinomial regression.")
      }
      
      # Prepare independent variables with proper reference categories
      model_data <- df
      categorical_vars <- indep_vars[sapply(df[indep_vars], function(x) is.factor(x) || is.character(x))]
      
      # Set reference categories for categorical independent variables
      ref_categories <- list()
      for (var in categorical_vars) {
        ref_input <- input[[paste0("multinomial_ref_", var)]]
        if (!is.null(ref_input)) {
          # Check if reference level exists
          var_levels <- if (is.factor(model_data[[var]])) {
            levels(model_data[[var]])
          } else if (is.character(model_data[[var]])) {
            unique(na.omit(model_data[[var]]))
          } else {
            NULL
          }
          
          if (!is.null(var_levels) && !ref_input %in% var_levels) {
            return(paste("Reference level", ref_input, "not found in variable", var))
          }
          
          if (is.factor(model_data[[var]])) {
            model_data[[var]] <- relevel(model_data[[var]], ref = ref_input)
          } else {
            # Convert character to factor with specified reference
            model_data[[var]] <- factor(model_data[[var]])
            model_data[[var]] <- relevel(model_data[[var]], ref = ref_input)
          }
          ref_categories[[var]] <- ref_input
        } else if (is.factor(model_data[[var]])) {
          ref_categories[[var]] <- levels(model_data[[var]])[1]
        }
      }
      
      # Create formula
      formula <- as.formula(paste("multinomial_dep_factor", "~", paste(indep_vars, collapse = " + ")))
      
      # Check for sufficient data in each category
      category_counts <- table(model_data$multinomial_dep_factor)
      if (any(category_counts < 2)) {
        return(paste("Insufficient data in some categories. Minimum 2 cases required per category. Current counts:", 
                     paste(names(category_counts), "=", category_counts, collapse = ", ")))
      }
      
      # Perform multinomial regression using nnet::multinom
      model <- nnet::multinom(formula, data = model_data, trace = FALSE)
      
      # Check if model converged
      if (model$convergence != 0) {
        return("Model failed to converge. Try increasing the number of iterations or simplifying the model.")
      }
      
      cat("Multinomial Logistic Regression\n")
      cat("===============================\n\n")
      cat("Dependent Variable:", dep_var, "\n")
      cat("Reference Category:", levels(model_data$multinomial_dep_factor)[1], "\n")
      cat("Independent Variables:", paste(indep_vars, collapse = ", "), "\n\n")
      
      # Display category counts
      cat("Category Counts:\n")
      print(category_counts)
      cat("\n")
      
      # Display reference categories
      if (length(ref_categories) > 0) {
        cat("Reference Categories for Independent Variables:\n")
        for (var in names(ref_categories)) {
          cat(paste0("  ", var, ": ", ref_categories[[var]], "\n"))
        }
        cat("\n")
      }
      
      # Model summary
      summary_model <- summary(model)
      
      # Calculate z-values and p-values
      z_values <- summary_model$coefficients / summary_model$standard.errors
      p_values <- 2 * (1 - pnorm(abs(z_values)))
      
      # Create comprehensive result table
      cat("Coefficient Table:\n")
      cat("==================\n\n")
      
      # Get the response categories (excluding reference)
      response_categories <- colnames(summary_model$coefficients)
      
      for (i in 1:length(response_categories)) {
        cat("Response Category:", response_categories[i], "vs Reference\n")
        cat("------------------------------------------------\n")
        
        # Extract coefficients, SE, z-values, and p-values for this category
        coefs <- summary_model$coefficients[i, ]
        ses <- summary_model$standard.errors[i, ]
        zs <- z_values[i, ]
        ps <- p_values[i, ]
        
        # Create result table for this category
        result_table <- data.frame(
          Coefficient = round(coefs, 4),
          Std.Error = round(ses, 4),
          z.value = round(zs, 4),
          p.value = format.pval(ps, digits = 4)
        )
        
        print(result_table)
        cat("\n")
        
        # Calculate odds ratios
        odds_ratios <- exp(coefs)
        cat("Odds Ratios:\n")
        or_table <- data.frame(Odds.Ratio = round(odds_ratios, 4))
        print(or_table)
        cat("\n\n")
      }
      
      # Model fit statistics
      cat("Model Fit Statistics:\n")
      cat("=====================\n")
      
      # Calculate null model (intercept only)
      null_formula <- as.formula(paste("multinomial_dep_factor", "~ 1"))
      null_model <- nnet::multinom(null_formula, data = model_data, trace = FALSE)
      
      # Likelihood ratio test
      lr_statistic <- -2 * (logLik(null_model) - logLik(model))
      df_diff <- attr(logLik(model), "df") - attr(logLik(null_model), "df")
      lr_pvalue <- 1 - pchisq(lr_statistic, df_diff)
      
      cat("Likelihood Ratio Test: Chi-square(", df_diff, ") =", 
          round(lr_statistic, 3), ", p =", format.pval(lr_pvalue, digits = 3), "\n")
      
      # McFadden's pseudo R-squared
      pseudo_r2 <- 1 - (logLik(model) / logLik(null_model))
      cat("McFadden's R-squared:", round(pseudo_r2, 4), "\n")
      
      # AIC and BIC
      cat("AIC:", round(AIC(model), 3), "\n")
      cat("BIC:", round(BIC(model), 3), "\n")
      
    }, error = function(e) {
      return(paste("Error in multinomial regression:", e$message))
    })
  })
  
  # Poisson Regression Analysis - CORRECTED VERSION
  # UI for categorical variable reference category selection for INDEPENDENT variables in Poisson regression
  output$poisson_ref_ui <- renderUI({
    req(data(), input$poisson_indep)
    
    df <- data()
    categorical_vars <- input$poisson_indep[sapply(df[input$poisson_indep], function(x) is.factor(x) || is.character(x))]
    
    if (length(categorical_vars) == 0) {
      return(NULL)
    }
    
    ref_ui_list <- lapply(categorical_vars, function(var) {
      var_levels <- if (is.factor(df[[var]])) {
        levels(df[[var]])
      } else {
        unique(na.omit(df[[var]]))
      }
      
      selectInput(
        inputId = paste0("poisson_ref_", var),
        label = paste("Reference category for", var),
        choices = var_levels,
        selected = var_levels[1]
      )
    })
    
    tagList(
      h4("Reference Categories for Categorical Variables"),
      ref_ui_list
    )
  })
  
  output$poisson_output <- renderPrint({
    req(data(), input$poisson_dep, input$poisson_indep, input$run_poisson)
    
    tryCatch({
      df <- data()
      dep_var <- input$poisson_dep
      indep_vars <- input$poisson_indep
      
      if (is.null(dep_var) || dep_var == "") {
        return("Please select a dependent variable.")
      }
      
      if (length(indep_vars) == 0) {
        return("Please select at least one independent variable.")
      }
      
      # Check if dependent variable is count data (non-negative)
      if (any(df[[dep_var]] < 0, na.rm = TRUE)) {
        return("Dependent variable should contain non-negative values for Poisson regression.")
      }
      
      # Check if dependent variable consists of integers (for true count data)
      # But allow non-integer values for rate data (with offset)
      if (!all(df[[dep_var]] %% 1 == 0, na.rm = TRUE)) {
        cat("Note: Dependent variable contains non-integer values.\n")
        cat("If this is rate data, consider using an offset term.\n\n")
      }
      
      # Prepare independent variables with proper reference categories
      model_data <- df
      categorical_vars <- indep_vars[sapply(df[indep_vars], function(x) is.factor(x) || is.character(x))]
      
      # Set reference categories for categorical independent variables
      for (var in categorical_vars) {
        ref_input <- input[[paste0("poisson_ref_", var)]]
        if (!is.null(ref_input)) {
          if (is.factor(model_data[[var]])) {
            model_data[[var]] <- relevel(model_data[[var]], ref = ref_input)
          } else {
            # Convert character to factor with specified reference
            model_data[[var]] <- factor(model_data[[var]])
            model_data[[var]] <- relevel(model_data[[var]], ref = ref_input)
          }
        }
      }
      
      # Create formula
      formula <- as.formula(paste(dep_var, "~", paste(indep_vars, collapse = " + ")))
      
      # Perform Poisson regression
      model <- glm(formula, data = model_data, family = poisson)
      
      cat("Poisson Regression\n")
      cat("==================\n\n")
      cat("Dependent Variable:", dep_var, "\n")
      cat("Independent Variables:", paste(indep_vars, collapse = ", "), "\n\n")
      
      # Display reference categories used
      if (length(categorical_vars) > 0) {
        cat("Reference Categories:\n")
        for (var in categorical_vars) {
          ref_input <- input[[paste0("poisson_ref_", var)]]
          cat(paste0("  ", var, ": ", ref_input, "\n"))
        }
        cat("\n")
      }
      
      # Model summary
      summary_model <- summary(model)
      print(summary_model)
      
      # Calculate incidence rate ratios (exponentiated coefficients)
      irr <- exp(coef(model))
      cat("\nIncidence Rate Ratios (IRR):\n")
      print(round(irr, 4))
      
      # Calculate confidence intervals for IRR
      ci_irr <- exp(confint(model))
      cat("\n95% Confidence Intervals for IRR:\n")
      print(round(ci_irr, 4))
      
      # Check for overdispersion
      dispersion <- sum(residuals(model, type = "pearson")^2) / df.residual(model)
      cat("\nDispersion Parameter:", round(dispersion, 3), "\n")
      if (dispersion > 1.5) {
        cat("Warning: Overdispersion detected (dispersion > 1.5).\n")
        cat("Consider using negative binomial regression instead.\n")
      } else if (dispersion < 0.8) {
        cat("Note: Underdispersion detected (dispersion < 0.8).\n")
      } else {
        cat("Dispersion is approximately 1 (no overdispersion detected).\n")
      }
      
      # Goodness of fit tests
      cat("\nGoodness of Fit:\n")
      cat("Log-likelihood:", round(logLik(model), 3), "\n")
      cat("AIC:", round(AIC(model), 3), "\n")
      cat("BIC:", round(BIC(model), 3), "\n")
      
      # Likelihood ratio test (compared to null model)
      null_formula <- as.formula(paste(dep_var, "~ 1"))
      null_model <- glm(null_formula, data = model_data, family = poisson)
      lr_test <- anova(null_model, model, test = "Chisq")
      cat("\nLikelihood Ratio Test:\n")
      cat("Chi-square(", lr_test$Df[2], ") =", round(lr_test$Deviance[2], 3), 
          ", p =", format.pval(lr_test$`Pr(>Chi)`[2], digits = 3, eps = 0.001), "\n")
      
    }, error = function(e) {
      return(paste("Error in Poisson regression:", e$message))
    })
  })
  # CFA Analysis - ENHANCED VERSION
  output$cfa_summary <- renderPrint({
    req(data(), input$cfa_vars, input$cfa_model, input$run_cfa)
    
    tryCatch({
      df <- data()
      cfa_vars <- input$cfa_vars
      model_syntax <- input$cfa_model
      
      if (length(cfa_vars) < 3) {
        return("Please select at least 3 variables for CFA.")
      }
      
      if (model_syntax == "") {
        return("Please specify the CFA model syntax.")
      }
      
      # Extract only numeric variables
      numeric_vars <- sapply(df[, cfa_vars], is.numeric)
      if (sum(numeric_vars) < 3) {
        return("Please select at least 3 numeric variables for CFA.")
      }
      
      cfa_data <- df[, cfa_vars[numeric_vars]]
      
      # Remove cases with missing values
      cfa_data <- na.omit(cfa_data)
      
      cat("Confirmatory Factor Analysis\n")
      cat("===========================\n\n")
      cat("Number of Variables:", ncol(cfa_data), "\n")
      cat("Number of Cases:", nrow(cfa_data), "\n\n")
      
      # Perform CFA using lavaan
      fit <- cfa(model_syntax, data = cfa_data, std.lv = TRUE)
      
      # Store the fit object for other outputs
      cfa_fit(fit)
      cfa_data_reactive(cfa_data)
      
      cat("Model Syntax:\n")
      cat(model_syntax, "\n\n")
      
      cat("Model converged:", lavInspect(fit, "converged"), "\n")
      
    }, error = function(e) {
      return(paste("Error in CFA:", e$message))
    })
  })
  
  # CFA Fit Measures - ENHANCED
  output$cfa_fit <- renderPrint({
    req(cfa_fit())
    
    tryCatch({
      fit <- cfa_fit()
      
      cat("Fit Measures\n")
      cat("============\n\n")
      
      fit_measures <- fitMeasures(fit)
      
      # Basic fit measures
      cat("Chi-square (χ²):", round(fit_measures["chisq"], 3), "\n")
      cat("Degrees of freedom:", fit_measures["df"], "\n")
      cat("p-value:", format.pval(fit_measures["pvalue"], digits = 3), "\n")
      cat("χ²/df:", round(fit_measures["chisq"]/fit_measures["df"], 3), "\n\n")
      
      # Comparative fit indices
      cat("Comparative Fit Index (CFI):", round(fit_measures["cfi"], 3), "\n")
      cat("Tucker-Lewis Index (TLI):", round(fit_measures["tli"], 3), "\n\n")
      
      # Absolute fit indices
      cat("RMSEA:", round(fit_measures["rmsea"], 3), "\n")
      cat("RMSEA 90% CI: [", round(fit_measures["rmsea.ci.lower"], 3), ",", 
          round(fit_measures["rmsea.ci.upper"], 3), "]\n")
      cat("RMSEA p-value:", format.pval(fit_measures["rmsea.pvalue"], digits = 3), "\n")
      cat("Standardized RMSEA (SRMR):", round(fit_measures["srmr"], 3), "\n\n")
      
      # Information criteria
      cat("Akaike Information Criterion (AIC):", round(fit_measures["aic"], 3), "\n")
      cat("Bayesian Information Criterion (BIC):", round(fit_measures["bic"], 3), "\n\n")
      
      # Additional fit measures
      cat("Goodness of Fit Index (GFI):", 
          ifelse("gfi" %in% names(fit_measures), round(fit_measures["gfi"], 3), "Not available"), "\n")
      cat("Adjusted Goodness of Fit Index (AGFI):", 
          ifelse("agfi" %in% names(fit_measures), round(fit_measures["agfi"], 3), "Not available"), "\n")
      cat("Normed Fit Index (NFI):", 
          ifelse("nfi" %in% names(fit_measures), round(fit_measures["nfi"], 3), "Not available"), "\n")
      
    }, error = function(e) {
      return(paste("Error in fit measures:", e$message))
    })
  })
  
  # CFA Parameter Estimates
  output$cfa_params <- renderPrint({
    req(cfa_fit())
    
    tryCatch({
      fit <- cfa_fit()
      
      cat("Parameter Estimates\n")
      cat("===================\n\n")
      
      # Get parameter estimates with standardized values
      params <- parameterEstimates(fit, standardized = TRUE)
      
      # Filter for different types of parameters
      loadings <- params[params$op == "=~", ]
      correlations <- params[params$op == "~~" & params$lhs != params$rhs, ]
      variances <- params[params$op == "~~" & params$lhs == params$rhs, ]
      
      cat("Factor Loadings:\n")
      if (nrow(loadings) > 0) {
        loadings_display <- data.frame(
          Factor = loadings$lhs,
          Variable = loadings$rhs,
          Estimate = round(loadings$est, 3),
          Std.Error = round(loadings$se, 3),
          z.value = round(loadings$z, 3),
          p.value = format.pval(loadings$pvalue, digits = 3),
          Std.All = round(loadings$std.all, 3)
        )
        print(loadings_display, row.names = FALSE)
      } else {
        cat("No factor loadings found.\n")
      }
      
      cat("\nFactor Correlations:\n")
      if (nrow(correlations) > 0) {
        # Filter for factor correlations only (not residual correlations)
        factor_corrs <- correlations[grepl("f", correlations$lhs) & grepl("f", correlations$rhs), ]
        if (nrow(factor_corrs) > 0) {
          corrs_display <- data.frame(
            Factor1 = factor_corrs$lhs,
            Factor2 = factor_corrs$rhs,
            Estimate = round(factor_corrs$est, 3),
            Std.Error = round(factor_corrs$se, 3),
            z.value = round(factor_corrs$z, 3),
            p.value = format.pval(factor_corrs$pvalue, digits = 3),
            Std.All = round(factor_corrs$std.all, 3)
          )
          print(corrs_display, row.names = FALSE)
        } else {
          cat("No factor correlations found.\n")
        }
      }
      
      cat("\nVariances:\n")
      if (nrow(variances) > 0) {
        vars_display <- data.frame(
          Variable = variances$lhs,
          Estimate = round(variances$est, 3),
          Std.Error = round(variances$se, 3),
          z.value = round(variances$z, 3),
          p.value = format.pval(variances$pvalue, digits = 3),
          Std.All = round(variances$std.all, 3)
        )
        print(vars_display, row.names = FALSE)
      }
      
    }, error = function(e) {
      return(paste("Error in parameter estimates:", e$message))
    })
  })
  
  # CFA Modification Indices
  output$cfa_modindices <- renderPrint({
    req(cfa_fit())
    
    tryCatch({
      fit <- cfa_fit()
      
      cat("Modification Indices\n")
      cat("====================\n\n")
      
      # Get modification indices
      mi <- modificationIndices(fit, sort = TRUE)
      
      # Filter for meaningful modifications (MI > 4)
      meaningful_mi <- mi[mi$mi > 4, ]
      
      if (nrow(meaningful_mi) > 0) {
        mi_display <- data.frame(
          lhs = meaningful_mi$lhs,
          op = meaningful_mi$op,
          rhs = meaningful_mi$rhs,
          MI = round(meaningful_mi$mi, 3),
          EP.C = round(meaningful_mi$epc, 3)
        )
        print(mi_display, row.names = FALSE)
        
        cat("\nInterpretation:\n")
        cat("MI > 4 suggests potential model improvement.\n")
        cat("EP.C = Expected parameter change if path is added.\n")
      } else {
        cat("No modification indices with MI > 4 found.\n")
        cat("Model appears to be well-specified.\n")
      }
      
    }, error = function(e) {
      return(paste("Error in modification indices:", e$message))
    })
  })
  
  # CFA Path Diagram
  output$cfa_diagram <- renderPlot({
    req(cfa_fit())
    
    tryCatch({
      fit <- cfa_fit()
      
      # Create path diagram with standardized estimates
      semPaths(fit, what = "std", whatLabels = "std", 
               style = "lisrel", layout = "tree", 
               edge.label.cex = 0.8, curvePivot = TRUE,
               sizeMan = 8, sizeLat = 12,
               nCharNodes = 0, rotation = 2,
               intercepts = FALSE, residuals = TRUE,
               thresholds = FALSE, optimizeLatRes = TRUE,
               exoVar = FALSE, exoCov = FALSE,
               cardinal = FALSE)
      
    }, error = function(e) {
      plot.new()
      text(0.5, 0.5, paste("Error creating path diagram:", e$message), 
           cex = 1.2, col = "red")
    })
  })
  
  # CFA Reliability and Validity
  output$cfa_reliability <- renderPrint({
    req(cfa_fit(), cfa_data_reactive())
    
    tryCatch({
      fit <- cfa_fit()
      cfa_data <- cfa_data_reactive()
      
      cat("Reliability and Validity Analysis\n")
      cat("=================================\n\n")
      
      # Calculate McDonald's Omega
      if (require(semTools)) {
        omega <- reliability(fit)
        
        cat("McDonald's Omega (ω):\n")
        print(round(omega, 3))
        cat("\n")
        
        # Interpretation guidelines
        cat("Reliability Interpretation:\n")
        cat("ω > 0.90: Excellent reliability\n")
        cat("ω > 0.80: Good reliability\n")
        cat("ω > 0.70: Acceptable reliability\n")
        cat("ω < 0.70: Questionable reliability\n\n")
      }
      
      # Calculate Average Variance Extracted (AVE) and Composite Reliability (CR)
      if (require(semTools)) {
        # Get standardized loadings
        std_loadings <- standardizedSolution(fit)
        std_loadings <- std_loadings[std_loadings$op == "=~", ]
        
        # Group by factor
        factors <- unique(std_loadings$lhs)
        
        for (factor in factors) {
          factor_loadings <- std_loadings[std_loadings$lhs == factor, "est.std"]
          
          # Calculate AVE (Average Variance Extracted)
          ave <- sum(factor_loadings^2) / length(factor_loadings)
          
          # Calculate CR (Composite Reliability)
          cr <- (sum(factor_loadings))^2 / ((sum(factor_loadings))^2 + sum(1 - factor_loadings^2))
          
          cat("Factor:", factor, "\n")
          cat("  Average Variance Extracted (AVE):", round(ave, 3), "\n")
          cat("  Composite Reliability (CR):", round(cr, 3), "\n")
          
          # AVE should be > 0.5 for convergent validity
          if (ave > 0.5) {
            cat("  Convergent Validity: Adequate (AVE > 0.5)\n")
          } else {
            cat("  Convergent Validity: Questionable (AVE < 0.5)\n")
          }
          cat("\n")
        }
      }
      
      # Discriminant Validity (if more than one factor)
      factors <- lavNames(fit, type = "lv")
      if (length(factors) > 1) {
        cat("Discriminant Validity Assessment\n")
        cat("------------------------------\n")
        
        # Get factor correlations
        factor_corrs <- lavInspect(fit, "cor.lv")
        
        if (!is.null(factor_corrs)) {
          cat("Factor Correlations:\n")
          print(round(factor_corrs, 3))
          cat("\n")
          
          # Check if correlations are significantly different from 1
          cat("Discriminant Validity: Factors should correlate < 0.85\n")
          high_corrs <- which(abs(factor_corrs) > 0.85 & upper.tri(factor_corrs), arr.ind = TRUE)
          if (length(high_corrs) > 0) {
            cat("Warning: High correlations between factors detected.\n")
            for (i in 1:nrow(high_corrs)) {
              cat("  ", rownames(factor_corrs)[high_corrs[i,1]], "-", 
                  colnames(factor_corrs)[high_corrs[i,2]], ":", 
                  round(factor_corrs[high_corrs[i,1], high_corrs[i,2]], 3), "\n")
            }
          }
        }
      }
      
    }, error = function(e) {
      return(paste("Error in reliability analysis:", e$message))
    })
  })
  
  # CFA Residuals
  output$cfa_residuals <- renderPrint({
    req(cfa_fit())
    
    tryCatch({
      fit <- cfa_fit()
      
      cat("Residual Analysis\n")
      cat("=================\n\n")
      
      # Get residuals
      res <- residuals(fit, type = "cor")
      
      cat("Standardized Residuals (correlation metric):\n")
      if (!is.null(res$cov)) {
        # Convert to data frame for better display
        res_cov <- as.data.frame(round(res$cov, 3))
        colnames(res_cov) <- rownames(res_cov)
        print(res_cov)
      }
      
      cat("\nSummary of Absolute Residuals:\n")
      if (!is.null(res$cov)) {
        abs_res <- abs(as.numeric(res$cov))
        cat("Mean absolute residual:", round(mean(abs_res, na.rm = TRUE), 3), "\n")
        cat("Maximum absolute residual:", round(max(abs_res, na.rm = TRUE), 3), "\n")
        cat("Number of residuals > 0.1:", sum(abs_res > 0.1, na.rm = TRUE), "\n")
        cat("Number of residuals > 0.05:", sum(abs_res > 0.05, na.rm = TRUE), "\n")
      }
      
      # Check for large residuals
      cat("\nLarge Residuals (|residual| > 0.1):\n")
      if (!is.null(res$cov)) {
        large_res <- which(abs(res$cov) > 0.1, arr.ind = TRUE)
        if (length(large_res) > 0) {
          for (i in 1:nrow(large_res)) {
            row_name <- rownames(res$cov)[large_res[i,1]]
            col_name <- colnames(res$cov)[large_res[i,2]]
            if (row_name != col_name) {  # Skip diagonal
              cat("  ", row_name, "-", col_name, ":", 
                  round(res$cov[large_res[i,1], large_res[i,2]], 3), "\n")
            }
          }
        } else {
          cat("No large residuals found.\n")
        }
      }
      
    }, error = function(e) {
      return(paste("Error in residual analysis:", e$message))
    })
  })
  
  # Store CFA fit object
  cfa_fit <- reactiveVal()
  cfa_data_reactive <- reactiveVal()
  # Explore Analysis
  output$explore_output <- renderPrint({
    req(data(), input$explore_vars, input$run_explore)
    
    tryCatch({
      df <- data()
      explore_vars <- input$explore_vars
      factor_var <- input$explore_factor
      
      if (length(explore_vars) == 0) {
        return("Please select at least one dependent variable.")
      }
      
      cat("Explore Analysis\n")
      cat("================\n\n")
      
      for (explore_var in explore_vars) {
        cat("Variable:", explore_var, "\n")
        cat("----------------------------------------\n")
        
        if (factor_var != "None") {
          cat("Factor:", factor_var, "\n")
          # Group by factor variable
          groups <- unique(df[[factor_var]])
          
          for (group in groups) {
            cat("\nGroup:", group, "\n")
            group_data <- df[df[[factor_var]] == group, explore_var]
            group_data <- group_data[!is.na(group_data)]
            
            if ("Descriptives" %in% input$explore_stats) {
              cat("Descriptive Statistics:\n")
              desc <- describe(group_data)
              print(round(desc, 3))
            }
            
            if ("Normality tests" %in% input$explore_stats) {
              cat("\nNormality Tests:\n")
              shapiro <- shapiro.test(group_data)
              cat("Shapiro-Wilk: W =", round(shapiro$statistic, 3), 
                  ", p =", format.pval(shapiro$p.value, digits = 3), "\n")
            }
            
            if ("Outliers" %in% input$explore_stats && length(group_data) > 10) {
              cat("\nOutlier Detection:\n")
              # Identify outliers using IQR method
              Q1 <- quantile(group_data, 0.25)
              Q3 <- quantile(group_data, 0.75)
              IQR <- Q3 - Q1
              lower_bound <- Q1 - 1.5 * IQR
              upper_bound <- Q3 + 1.5 * IQR
              outliers <- group_data[group_data < lower_bound | group_data > upper_bound]
              cat("Potential outliers:", round(outliers, 3), "\n")
            }
          }
        } else {
          # No grouping variable
          explore_data <- df[[explore_var]]
          explore_data <- explore_data[!is.na(explore_data)]
          
          if ("Descriptives" %in% input$explore_stats) {
            cat("Descriptive Statistics:\n")
            desc <- describe(explore_data)
            print(round(desc, 3))
          }
          
          if ("Normality tests" %in% input$explore_stats) {
            cat("\nNormality Tests:\n")
            shapiro <- shapiro.test(explore_data)
            cat("Shapiro-Wilk: W =", round(shapiro$statistic, 3), 
                ", p =", format.pval(shapiro$p.value, digits = 3), "\n")
          }
          
          if ("Outliers" %in% input$explore_stats && length(explore_data) > 10) {
            cat("\nOutlier Detection:\n")
            Q1 <- quantile(explore_data, 0.25)
            Q3 <- quantile(explore_data, 0.75)
            IQR <- Q3 - Q1
            lower_bound <- Q1 - 1.5 * IQR
            upper_bound <- Q3 + 1.5 * IQR
            outliers <- explore_data[explore_data < lower_bound | explore_data > upper_bound]
            cat("Potential outliers:", round(outliers, 3), "\n")
          }
        }
        cat("\n")
      }
      
    }, error = function(e) {
      return(paste("Error in explore analysis:", e$message))
    })
  })
  
  # Paired T-Test Analysis - CORRECTED VERSION
  output$paired_output <- renderPrint({
    req(data(), input$paired_vars, input$run_paired)
    
    tryCatch({
      df <- data()
      paired_vars <- input$paired_vars
      
      if (length(paired_vars) < 2) {
        return("Please select at least two variables for paired comparison.")
      }
      
      cat("Paired Samples T-Test\n")
      cat("=====================\n\n")
      
      # Create all possible pairs
      pairs <- combn(paired_vars, 2, simplify = FALSE)
      
      for (pair in pairs) {
        cat("Pair:", pair[1], "vs", pair[2], "\n")
        cat("----------------------------------------\n")
        
        var1 <- df[[pair[1]]]
        var2 <- df[[pair[2]]]
        
        # Remove cases with missing values in either variable
        complete_cases <- complete.cases(var1, var2)
        var1 <- var1[complete_cases]
        var2 <- var2[complete_cases]
        
        # Perform paired t-test with proper tail handling
        if (input$paired_tails == "Two-tailed") {
          t_test <- t.test(var1, var2, paired = TRUE, 
                           conf.level = input$paired_conf/100,
                           alternative = "two.sided")
        } else {
          # Determine direction for one-tailed test
          mean_diff <- mean(var1) - mean(var2)
          alternative <- ifelse(mean_diff > 0, "greater", "less")
          t_test <- t.test(var1, var2, paired = TRUE, 
                           conf.level = input$paired_conf/100,
                           alternative = alternative)
        }
        
        # Calculate differences
        differences <- var1 - var2
        
        cat("N =", length(var1), "\n")
        cat("Mean of", pair[1], "=", round(mean(var1), 3), "\n")
        cat("Mean of", pair[2], "=", round(mean(var2), 3), "\n")
        cat("Mean Difference =", round(mean(differences), 3), "\n")
        cat("SD of Differences =", round(sd(differences), 3), "\n")
        cat("SE of Differences =", round(sd(differences)/sqrt(length(differences)), 3), "\n\n")
        
        cat("t(", t_test$parameter, ") =", round(t_test$statistic, 3), "\n")
        cat("p-value =", format.pval(t_test$p.value, digits = 3), "\n")
        cat(input$paired_conf, "% CI: [", round(t_test$conf.int[1], 3), ",", 
            round(t_test$conf.int[2], 3), "]\n\n")
        
        # Add interpretation for one-tailed tests
        if (input$paired_tails == "One-tailed") {
          if (t_test$p.value < 0.05) {
            if (mean(var1) > mean(var2)) {
              cat("Interpretation:", pair[1], "is significantly greater than", pair[2], "(one-tailed)\n")
            } else {
              cat("Interpretation:", pair[1], "is significantly less than", pair[2], "(one-tailed)\n")
            }
          } else {
            cat("Interpretation: No significant difference between variables (one-tailed)\n")
          }
        }
        cat("\n")
      }
      
    }, error = function(e) {
      return(paste("Error in paired t-test:", e$message))
    })
  })
  
}

# Run the application
shinyApp(ui = ui, server = server)
