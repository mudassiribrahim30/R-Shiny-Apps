library(shiny)
library(shinydashboard)
library(ggplot2)
library(dplyr)
library(ggpubr)
library(survival)
library(survminer)
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
library(fresh)
library(lm.beta)
library(MASS)
library(nnet)
library(dunn.test)
library(openxlsx)
library(janitor)
library(lubridate)
library(shinycssloaders)
library(shinyjs)
library(colourpicker)
library(flextable)

# Custom CSS for better layout and animations
custom_css <- "
.sidebar {
  overflow-y: auto;
  max-height: 90vh;
  position: sticky;
  top: 0;
}

.main-content {
  overflow-y: auto;
  max-height: 90vh;
}

.navbar {
  position: sticky;
  top: 0;
  z-index: 1000;
}

.jumbotron {
  background: linear-gradient(rgba(255,255,255,0.9), rgba(255,255,255,0.9)), 
              url('');
  background-size: cover;
  padding: 3rem 2rem;
  margin-bottom: 2rem;
  border-radius: 0.3rem;
}

.card {
  margin-bottom: 1rem;
  border: 1px solid #e9ecef;
  border-radius: 0.25rem;
  transition: all 0.3s ease;
}

.card:hover {
  transform: translateY(-5px);
  box-shadow: 0 8px 25px rgba(0,0,0,0.15);
}

.card-body {
  padding: 1.25rem;
}

/* Loading animations */
.loading-container {
  display: flex;
  justify-content: center;
  align-items: center;
  padding: 2rem;
}

.loading-spinner {
  border: 4px solid #f3f3f3;
  border-top: 4px solid #0C5EA8;
  border-radius: 50%;
  width: 40px;
  height: 40px;
  animation: spin 2s linear infinite;
}

@keyframes spin {
  0% { transform: rotate(0deg); }
  100% { transform: rotate(360deg); }
}

/* Interactive button styles */
.btn-action {
  transition: all 0.3s ease;
  position: relative;
  overflow: hidden;
}

.btn-action:hover {
  transform: translateY(-2px);
  box-shadow: 0 4px 12px rgba(0,0,0,0.15);
}

.btn-action:active {
  transform: translateY(0);
}

.btn-action::after {
  content: '';
  position: absolute;
  top: 50%;
  left: 50%;
  width: 0;
  height: 0;
  border-radius: 50%;
  background: rgba(255,255,255,0.3);
  transform: translate(-50%, -50%);
  transition: width 0.3s, height 0.3s;
}

.btn-action:active::after {
  width: 100px;
  height: 100px;
}

/* Sidebar and main content layout */
.sidebar-panel {
  background: #f8f9fa;
  border-right: 1px solid #dee2e6;
  padding: 20px;
  height: calc(100vh - 80px);
  overflow-y: auto;
  position: sticky;
  top: 0;
}

.main-panel {
  padding: 20px;
  height: calc(100vh - 80px);
  overflow-y: auto;
}

/* Result card animations */
.result-card {
  transition: all 0.3s ease;
  border-left: 4px solid #0C5EA8;
}

.result-card:hover {
  box-shadow: 0 4px 15px rgba(0,0,0,0.1);
  transform: translateX(5px);
}

/* Tab animations */
.tab-pane {
  animation: fadeIn 0.5s ease-in;
}

@keyframes fadeIn {
  from { opacity: 0; transform: translateY(10px); }
  to { opacity: 1; transform: translateY(0); }
}

/* Custom scrollbar */
::-webkit-scrollbar {
  width: 8px;
}

::-webkit-scrollbar-track {
  background: #f1f1f1;
  border-radius: 10px;
}

::-webkit-scrollbar-thumb {
  background: #0C5EA8;
  border-radius: 10px;
}

::-webkit-scrollbar-thumb:hover {
  background: #094580;
}

/* Pulse animation for important elements */
.pulse {
  animation: pulse 2s infinite;
}

@keyframes pulse {
  0% { box-shadow: 0 0 0 0 rgba(12, 94, 168, 0.4); }
  70% { box-shadow: 0 0 0 10px rgba(12, 94, 168, 0); }
  100% { box-shadow: 0 0 0 0 rgba(12, 94, 168, 0); }
}

/* Success and error states */
.alert-success {
  border-left: 4px solid #28a745;
}

.alert-danger {
  border-left: 4px solid #dc3545;
}

.alert-warning {
  border-left: 4px solid #ffc107;
}
"

# Helper function for survival object creation
create_surv_obj <- function(time_var, event_var, data) {
  time <- data[[time_var]]
  event <- data[[event_var]]
  
  # Convert event to numeric if it's not
  if (!is.numeric(event)) {
    event <- as.numeric(as.factor(event)) - 1
  }
  
  Surv(time = time, event = event)
}

# Helper function to handle labelled data
convert_labelled <- function(df) {
  df <- df %>%
    mutate(across(where(haven::is.labelled), as_factor)) %>%
    mutate(across(where(is.factor), as.character)) %>%
    mutate(across(where(is.character), ~ifelse(. == "", NA, .))) %>%
    mutate(across(everything(), ~ifelse(is.na(.), NA, .)))
  return(df)
}

# UI
ui <- navbarPage(
  title = div(
    img(src = "", 
        height = "30", 
        style = "margin-right:10px;padding-bottom:3px;"),
    "📈 EpiDem Suite™"
  ),
  footer = div(
    style = "text-align: center; padding: 10px; background-color: #f8f9fa; border-top: 1px solid #dee2e6;",
    "Copyright (c) Mudasir Mohammed Ibrahim (mudassiribrahim30@gmail.com)"
  ),
  windowTitle = "EpiDem Suite - Epidemiological Analysis Tool",
  id = "nav",
  header = tags$head(
    tags$style(HTML(custom_css)),
    useShinyjs()
  ),
  theme = bslib::bs_theme(version = 5, 
                          bootswatch = "flatly",
                          primary = "#0C5EA8",   # CDC Blue
                          secondary = "#CD2026", # CDC Red
                          success = "#5CB85C",
                          font_scale = 0.95),
  
  # Home/About Tab
  tabPanel("Home", icon = icon("house"),
           div(class = "container-fluid",
               div(class = "jumbotron",
                   h1("EpiDem Suite", class = "display-4"),
                   p("A comprehensive epidemiological analysis platform for public health professionals.", 
                     class = "lead"),
                   hr(class = "my-4"),
                   p("Upload your dataset and utilize the tools in the navigation bar to perform descriptive analyses, 
                     calculate disease frequency measures, assess statistical associations, and build regression models."),
                   p("Supported file formats: CSV, Excel (.xls, .xlsx), Stata (.dta), SPSS (.sav), SAS (.sas7bdat), RData (.RData, .rda)"),
                   p("This tool follows established guidelines for epidemiological analysis. Developed by Mudasir Mohammed Ibrahim (BSc, RN)"),
                   actionButton("goto_data", "Get Started →", 
                                class = "btn-primary btn-lg btn-action", 
                                icon = icon("rocket"))
               ),
               
               fluidRow(
                 column(4,
                        div(class = "card pulse",
                            div(class = "card-body",
                                h3("📊 Descriptive Analysis", class = "card-title"),
                                p("Explore frequency distributions, summary statistics, and create epidemic curves to understand your data's basic characteristics."),
                                tags$ul(
                                  tags$li("Frequency distributions and bar plots"),
                                  tags$li("Central tendency measures and histograms"),
                                  tags$li("Epidemic curves with customizable intervals"),
                                  tags$li("Spatial analysis and demographic breakdowns")
                                ),
                                actionButton("goto_descriptive", "Explore Descriptive Tools", 
                                             class = "btn-outline-primary btn-action")
                            )
                        )
                 ),
                 column(4,
                        div(class = "card pulse",
                            div(class = "card-body",
                                h3("🦠 Disease Metrics", class = "card-title"),
                                p("Calculate incidence rates, prevalence, attack rates, and mortality measures to quantify disease burden in populations."),
                                tags$ul(
                                  tags$li("Incidence rates and cumulative incidence"),
                                  tags$li("Prevalence calculations"),
                                  tags$li("Attack rates with confidence intervals"),
                                  tags$li("Mortality rates and case fatality rates"),
                                  tags$li("Life table analysis")
                                ),
                                actionButton("goto_disease", "Calculate Disease Metrics", 
                                             class = "btn-outline-primary btn-action")
                            )
                        )
                 ),
                 column(4,
                        div(class = "card pulse",
                            div(class = "card-body",
                                h3("📈 Statistical Tools", class = "card-title"),
                                p("Perform association tests, regression analyses, and survival analysis to identify risk factors and measure effects."),
                                tags$ul(
                                  tags$li("Risk ratios, odds ratios, and attributable risk"),
                                  tags$li("Poisson regression for count data"),
                                  tags$li("Survival analysis with Kaplan-Meier curves"),
                                  tags$li("Cox proportional hazards models"),
                                  tags$li("Linear and logistic regression"),
                                  tags$li("Chi-square tests with cross-tabulations")
                                ),
                                actionButton("goto_statistics", "Use Statistical Tools", 
                                             class = "btn-outline-primary btn-action")
                            )
                        )
                 )
               )
           )
  ),
  
  # Data Upload Tab
  tabPanel("Data Upload", icon = icon("database"),
           fluidRow(
             column(4, class = "sidebar-panel",
                    fileInput("file1", "Choose File",
                              accept = c(".csv", ".xls", ".xlsx", ".sav", ".dta", ".sas7bdat", ".RData", ".rda"),
                              multiple = FALSE),
                    helpText("Supported formats: CSV, Excel, Stata, SPSS, SAS, RData"),
                    tags$hr(),
                    conditionalPanel(
                      condition = "input.file1 && input.file1.type.includes('csv')",
                      checkboxInput("header", "Header", TRUE),
                      selectInput("sep", "Separator",
                                  choices = c(Comma = ",", Semicolon = ";", Tab = "\t", Space = " "),
                                  selected = ","),
                      selectInput("quote", "Quote",
                                  choices = c(None = "", "Double Quote" = '"', "Single Quote" = "'"),
                                  selected = '"')
                    ),
                    conditionalPanel(
                      condition = "input.file1 && (input.file1.type.includes('xls') || input.file1.type.includes('xlsx'))",
                      numericInput("sheet", "Sheet Number", value = 1, min = 1)
                    ),
                    actionButton("load_data", "Load Data", 
                                 class = "btn-primary btn-action",
                                 icon = icon("upload")),
                    tags$hr(),
                    uiOutput("missing_values_alert")
             ),
             column(8, class = "main-panel",
                    withSpinner(DTOutput("contents"), type = 4, color = "#0C5EA8"),
                    withSpinner(verbatimTextOutput("data_summary"), type = 4, color = "#0C5EA8")
             )
           )
  ),
  
  # Descriptive Analysis Tab - CORRECTED DOWNLOAD HANDLERS
  tabPanel("Descriptive Analysis", icon = icon("chart-bar"),
           tabsetPanel(
             tabPanel("Frequency Analysis",
                      fluidRow(
                        column(4, class = "sidebar-panel",
                               selectInput("freq_var", "Select Variable:", choices = NULL),
                               colourpicker::colourInput("freq_color", "Bar Color:", value = "#0C5EA8"),
                               textInput("freq_title", "Plot Title:", value = "Frequency Distribution"),
                               numericInput("freq_title_size", "Title Size:", value = 16, min = 8, max = 24),
                               numericInput("freq_x_size", "X-axis Label Size:", value = 14, min = 8, max = 20),
                               numericInput("freq_y_size", "Y-axis Label Size:", value = 14, min = 8, max = 20),
                               radioButtons("freq_type", "Download Format:",
                                            choices = c("Table", "Plot"), selected = "Table"),
                               actionButton("run_freq", "Run Analysis", 
                                            class = "btn-primary btn-action",
                                            icon = icon("play")),
                               downloadButton("download_freq", "Download Results",
                                              class = "btn-success btn-action")
                        ),
                        column(8, class = "main-panel",
                               withSpinner(DTOutput("freq_table"), type = 4, color = "#0C5EA8"),
                               withSpinner(plotlyOutput("freq_plot"), type = 4, color = "#0C5EA8")
                        )
                      )
             ),
             tabPanel("Central Tendency",
                      fluidRow(
                        column(4, class = "sidebar-panel",
                               selectInput("central_var", "Select Numeric Variable:", choices = NULL),
                               selectInput("central_stratify", "Stratify by (Optional):", 
                                           choices = c("None" = "none")),
                               colourpicker::colourInput("central_color", "Plot Color:", value = "#0C5EA8"),
                               textInput("central_title", "Plot Title:", value = "Central Tendency"),
                               numericInput("central_bins", "Number of Bins:", value = 30, min = 5, max = 100),
                               numericInput("central_title_size", "Title Size:", value = 16, min = 8, max = 24),
                               numericInput("central_x_size", "X-axis Label Size:", value = 14, min = 8, max = 20),
                               numericInput("central_y_size", "Y-axis Label Size:", value = 14, min = 8, max = 20),
                               radioButtons("central_type", "Download Format:",
                                            choices = c("Table", "Boxplot", "Histogram"), selected = "Table"),
                               actionButton("run_central", "Run Analysis", 
                                            class = "btn-primary btn-action",
                                            icon = icon("play")),
                               downloadButton("download_central", "Download Results",
                                              class = "btn-success btn-action")
                        ),
                        column(8, class = "main-panel",
                               withSpinner(DTOutput("central_table"), type = 4, color = "#0C5EA8"),
                               withSpinner(plotlyOutput("central_boxplot"), type = 4, color = "#0C5EA8"),
                               withSpinner(plotlyOutput("central_histogram"), type = 4, color = "#0C5EA8")
                        )
                      )
             ),
             tabPanel("Epidemic Curve",
                      fluidRow(
                        column(4, class = "sidebar-panel",
                               selectInput("epi_date", "Date Variable:", choices = NULL),
                               selectInput("epi_case", "Case Count Variable:", choices = NULL),
                               selectInput("epi_group", "Grouping Variable (Optional):", choices = c("None" = "none")),
                               selectInput("epi_interval", "Time Interval:",
                                           choices = c("Day", "Week", "Month", "Year"), selected = "Week"),
                               colourpicker::colourInput("epi_color", "Bar Color:", value = "#CD2026"),
                               textInput("epi_title", "Plot Title:", value = "Epidemic Curve"),
                               numericInput("epi_title_size", "Title Size:", value = 16, min = 8, max = 24),
                               numericInput("epi_x_size", "X-axis Label Size:", value = 14, min = 8, max = 20),
                               numericInput("epi_y_size", "Y-axis Label Size:", value = 14, min = 8, max = 20),
                               actionButton("run_epi", "Run Analysis", 
                                            class = "btn-primary btn-action",
                                            icon = icon("play")),
                               downloadButton("download_epicurve", "Download Plot",
                                              class = "btn-success btn-action"),
                               downloadButton("download_epi_data", "Download Data",
                                              class = "btn-info btn-action")
                        ),
                        column(8, class = "main-panel",
                               withSpinner(uiOutput("epi_summary"), type = 4, color = "#0C5EA8"),
                               withSpinner(plotlyOutput("epi_curve"), type = 4, color = "#0C5EA8")
                        )
                      )
             )
           )
  ),
  
  # Disease Frequency Tab
  tabPanel("Disease Frequency", icon = icon("viruses"),
           tabsetPanel(
             tabPanel("Incidence/Prevalence",
                      fluidRow(
                        column(4, class = "sidebar-panel",
                               selectInput("ip_measure", "Select Measure:",
                                           choices = c("Incidence Rate", "Cumulative Incidence", "Prevalence"),
                                           selected = "Incidence Rate"),
                               selectInput("ip_case", "Case Count Variable:", choices = NULL),
                               selectInput("ip_pop", "Population Variable:", choices = NULL),
                               conditionalPanel(
                                 condition = "input.ip_measure == 'Incidence Rate'",
                                 selectInput("ip_time", "Person-Time Variable:", choices = NULL)
                               ),
                               actionButton("run_ip", "Calculate", 
                                            class = "btn-primary btn-action",
                                            icon = icon("calculator")),
                               downloadButton("download_ip", "Download Results",
                                              class = "btn-success btn-action")
                        ),
                        column(8, class = "main-panel",
                               withSpinner(DTOutput("ip_table"), type = 4, color = "#0C5EA8")
                        )
                      )
             )
           )
  ),
  
  # Statistical Associations Tab - CORRECTED RISK RATIO SECTION
  tabPanel("Statistical Associations", icon = icon("calculator"),
           tabsetPanel(
             tabPanel("Risk Ratios",
                      fluidRow(
                        column(4, class = "sidebar-panel",
                               selectInput("rr_outcome", "Outcome Variable:", choices = NULL),
                               selectInput("rr_exposure", "Exposure Variable:", choices = NULL),
                               uiOutput("rr_outcome_levels_ui"),
                               uiOutput("rr_exposure_levels_ui"),
                               sliderInput("rr_conf", "Confidence Level:", min = 0.90, max = 0.99, value = 0.95, step = 0.01),
                               actionButton("run_rr", "Calculate", 
                                            class = "btn-primary btn-action",
                                            icon = icon("calculator")),
                               downloadButton("download_rr", "Download Results",
                                              class = "btn-success btn-action")
                        ),
                        column(8, class = "main-panel",
                               withSpinner(verbatimTextOutput("rr_results"), type = 4, color = "#0C5EA8"),
                               withSpinner(DTOutput("rr_table"), type = 4, color = "#0C5EA8"),
                               withSpinner(verbatimTextOutput("rr_interpretation"), type = 4, color = "#0C5EA8")
                        )
                      )
             ),
             
             tabPanel("Odds Ratios",
                      fluidRow(
                        column(4, class = "sidebar-panel",
                               selectInput("or_outcome", "Outcome Variable:", choices = NULL),
                               selectInput("or_exposure", "Exposure Variable:", choices = NULL),
                               sliderInput("or_conf", "Confidence Level:", min = 0.90, max = 0.99, value = 0.95, step = 0.01),
                               actionButton("run_or", "Calculate", 
                                            class = "btn-primary btn-action",
                                            icon = icon("calculator")),
                               downloadButton("download_or", "Download Results",
                                              class = "btn-success btn-action")
                        ),
                        column(8, class = "main-panel",
                               withSpinner(verbatimTextOutput("or_results"), type = 4, color = "#0C5EA8"),
                               withSpinner(DTOutput("or_table"), type = 4, color = "#0C5EA8")
                        )
                      )
             )
           )
  ),
  
  # Survival Analysis Tab
  tabPanel("Survival Analysis", icon = icon("heartbeat"),
           tabsetPanel(
             tabPanel("Kaplan-Meier",
                      tabsetPanel(
                        tabPanel("Analysis Settings",
                                 fluidRow(
                                   column(4, class = "sidebar-panel",
                                          selectInput("surv_time", "Time Variable:", choices = NULL),
                                          selectInput("surv_event", "Event Variable:", choices = NULL),
                                          selectInput("surv_group", "Grouping Variable (Optional):", choices = c("None" = "none")),
                                          actionButton("run_surv", "Run Analysis", 
                                                       class = "btn-primary btn-action",
                                                       icon = icon("play")),
                                          downloadButton("", "",
                                                         class = "btn-info btn-action")
                                   ),
                                   column(8, class = "main-panel",
                                          withSpinner(uiOutput("surv_stats"), type = 4, color = "#0C5EA8"),
                                          withSpinner(verbatimTextOutput("surv_summary"), type = 4, color = "#0C5EA8")
                                   )
                                 )
                        ),
                        tabPanel("Plot Options",
                                 fluidRow(
                                   column(4, class = "sidebar-panel",
                                          h4("Display Options"),
                                          checkboxInput("surv_ci", "Show Confidence Intervals", TRUE),
                                          checkboxInput("surv_risktable", "Show Risk Table", TRUE),
                                          checkboxInput("surv_median", "Show Median Survival Lines", TRUE),
                                          checkboxInput("surv_labels", "Show Survival Probability Labels", TRUE),
                                          br(),
                                          h4("Color Settings"),
                                          colourpicker::colourInput("surv_color1", "Group 1 Color:", value = "#0C5EA8"),
                                          colourpicker::colourInput("surv_color2", "Group 2 Color:", value = "#CD2026"),
                                          colourpicker::colourInput("surv_color3", "Group 3 Color:", value = "#5CB85C"),
                                          br(),
                                          h4("Text Settings"),
                                          textInput("surv_title", "Plot Title:", value = "Kaplan-Meier Curve"),
                                          textInput("surv_xlab", "X-axis Label:", value = "Time"),
                                          textInput("surv_ylab", "Y-axis Label:", value = "Survival Probability"),
                                          numericInput("surv_title_size", "Title Size:", value = 18, min = 8, max = 24),
                                          numericInput("surv_x_size", "X-axis Label Size:", value = 14, min = 8, max = 20),
                                          numericInput("surv_y_size", "Y-axis Label Size:", value = 14, min = 8, max = 20),
                                          numericInput("surv_label_size", "Label Size:", value = 5, min = 3, max = 10),
                                          br(),
                                          downloadButton("download_surv", "Download Plot",
                                                         class = "btn-success btn-action")
                                   ),
                                   column(8, class = "main-panel",
                                          withSpinner(plotOutput("surv_plot", height = "600px"), type = 4, color = "#0C5EA8")
                                   )
                                 )
                        ),
                        tabPanel("Life Table",
                                 fluidRow(
                                   column(4, class = "sidebar-panel",
                                          h4("Time Points Configuration"),
                                          helpText("Specify time points (comma-separated) to evaluate survival probabilities:"),
                                          textInput("life_table_times", "Time Points:", value = "0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0"),
                                          helpText("Example: 0.1, 0.25, 0.5, 0.75, 1.0 or specific time values like 30, 60, 90"),
                                          numericInput("life_table_decimals", "Decimal Places:", value = 4, min = 2, max = 6),
                                          actionButton("update_life_table", "Update Life Table", 
                                                       class = "btn-primary btn-action",
                                                       icon = icon("refresh")),
                                          br(), br(),
                                          downloadButton("download_life_table", "Download Life Table (Excel)",
                                                         class = "btn-success btn-action")
                                   ),
                                   column(8, class = "main-panel",
                                          h3("Life Table - Survival Probabilities at Specified Time Points"),
                                          helpText("This table shows the survival probabilities at the specified time points for each group."),
                                          withSpinner(DTOutput("life_table_output"), type = 4, color = "#0C5EA8"),
                                          br(),
                                          h4("Life Table Summary"),
                                          withSpinner(verbatimTextOutput("life_table_summary"), type = 4, color = "#0C5EA8")
                                   )
                                 )
                        )
                      )
             ),
             tabPanel("Cox Regression",
                      fluidRow(
                        column(4, class = "sidebar-panel",
                               selectInput("cox_time", "Time Variable:", choices = NULL),
                               selectInput("cox_event", "Event Variable:", choices = NULL),
                               selectizeInput("cox_vars", "Predictor Variables:", choices = NULL, multiple = TRUE),
                               uiOutput("cox_ref_ui"),
                               sliderInput("cox_conf", "Confidence Level:", min = 0.90, max = 0.99, value = 0.95, step = 0.01),
                               actionButton("run_cox", "Run Analysis", 
                                            class = "btn-primary btn-action",
                                            icon = icon("play")),
                               downloadButton("download_cox", "Download Results",
                                              class = "btn-success btn-action")
                        ),
                        column(8, class = "main-panel",
                               withSpinner(verbatimTextOutput("cox_summary"), type = 4, color = "#0C5EA8"),
                               withSpinner(DTOutput("cox_table"), type = 4, color = "#0C5EA8")
                        )
                      )
             ),
             # NEW POISSON REGRESSION SUBTAB
             tabPanel("Poisson Regression",
                      fluidRow(
                        column(4, class = "sidebar-panel",
                               selectInput("poisson_outcome", "Outcome Variable (Count):", choices = NULL),
                               selectizeInput("poisson_vars", "Predictor Variables:", choices = NULL, multiple = TRUE),
                               uiOutput("poisson_ref_ui"),
                               checkboxInput("poisson_offset", "Include Offset Variable", FALSE),
                               conditionalPanel(
                                 condition = "input.poisson_offset",
                                 selectInput("poisson_offset_var", "Offset Variable:", choices = NULL)
                               ),
                               sliderInput("poisson_conf", "Confidence Level:", min = 0.90, max = 0.99, value = 0.95, step = 0.01),
                               actionButton("run_poisson", "Run Analysis", 
                                            class = "btn-primary btn-action",
                                            icon = icon("play")),
                               downloadButton("download_poisson", "Download Results",
                                              class = "btn-success btn-action")
                        ),
                        column(8, class = "main-panel",
                               withSpinner(verbatimTextOutput("poisson_summary"), type = 4, color = "#0C5EA8"),
                               withSpinner(DTOutput("poisson_table"), type = 4, color = "#0C5EA8")
                        )
                      )
             )
           )
  ),
  
  # Regression Analysis Tab
  tabPanel("Regression Analysis", icon = icon("line-chart"),
           tabsetPanel(
             tabPanel("Linear Regression",
                      fluidRow(
                        column(4, class = "sidebar-panel",
                               selectInput("linear_outcome", "Outcome Variable:", choices = NULL),
                               selectizeInput("linear_vars", "Predictor Variables:", choices = NULL, multiple = TRUE),
                               uiOutput("linear_ref_ui"),
                               checkboxInput("linear_std", "Show Standardized Coefficients", FALSE),
                               actionButton("run_linear", "Run Analysis", 
                                            class = "btn-primary btn-action",
                                            icon = icon("play")),
                               downloadButton("download_linear", "Download Results",
                                              class = "btn-success btn-action")
                        ),
                        column(8, class = "main-panel",
                               withSpinner(verbatimTextOutput("linear_summary"), type = 4, color = "#0C5EA8"),
                               withSpinner(DTOutput("linear_table"), type = 4, color = "#0C5EA8")
                        )
                      )
             ),
             tabPanel("Logistic Regression", icon = icon("line-chart"),
                      fluidRow(
                        column(4, class = "sidebar-panel",
                               selectInput("logistic_outcome", "Outcome Variable:", choices = NULL),
                               uiOutput("logistic_outcome_ui"),
                               selectizeInput("logistic_vars", "Predictor Variables:", choices = NULL, multiple = TRUE),
                               uiOutput("logistic_ref_ui"),
                               checkboxInput("logistic_or", "Show Odds Ratios", TRUE),
                               actionButton("run_logistic", "Run Analysis", 
                                            class = "btn-primary btn-action",
                                            icon = icon("play")),
                               downloadButton("download_logistic", "Download Results",
                                              class = "btn-success btn-action")
                        ),
                        column(8, class = "main-panel",
                               withSpinner(verbatimTextOutput("logistic_summary"), type = 4, color = "#0C5EA8"),
                               withSpinner(DTOutput("logistic_table"), type = 4, color = "#0C5EA8")
                        )
                      )
             )
           )
  ),
  
  # Statistical Tests Tab - CORRECTED WITH ANOVA AND T-TESTS
  tabPanel("Statistical Tests", icon = icon("check-square"),
           tabsetPanel(
             tabPanel("Chi-square Test",
                      fluidRow(
                        column(4, class = "sidebar-panel",
                               selectInput("chisq_var1", "Variable 1:", choices = NULL),
                               selectInput("chisq_var2", "Variable 2:", choices = NULL),
                               checkboxInput("chisq_expected", "Show Expected Counts", FALSE),
                               checkboxInput("chisq_residuals", "Show Residuals", FALSE),
                               checkboxInput("chisq_fisher", "Include Fisher's Exact Test", FALSE),
                               checkboxInput("chisq_rr", "Include Relative Risk", FALSE),
                               actionButton("run_chisq", "Run Test", 
                                            class = "btn-primary btn-action",
                                            icon = icon("play")),
                               downloadButton("download_chisq", "Download Results",
                                              class = "btn-success btn-action")
                        ),
                        column(8, class = "main-panel",
                               withSpinner(verbatimTextOutput("chisq_results"), type = 4, color = "#0C5EA8"),
                               withSpinner(DTOutput("chisq_table"), type = 4, color = "#0C5EA8"),
                               withSpinner(plotOutput("chisq_plot"), type = 4, color = "#0C5EA8"),
                               downloadButton("download_chisq_plot", "Download Plot",
                                              class = "btn-info btn-action")
                        )
                      )
             ),
             
             # NEW: ANOVA SUBTAB
             tabPanel("ANOVA",
                      tabsetPanel(
                        tabPanel("One-Way ANOVA",
                                 fluidRow(
                                   column(4, class = "sidebar-panel",
                                          selectInput("anova_outcome", "Outcome Variable (Continuous):", choices = NULL),
                                          selectInput("anova_group", "Grouping Variable (Categorical):", choices = NULL),
                                          checkboxInput("anova_assumptions", "Check ANOVA Assumptions", TRUE),
                                          checkboxInput("anova_posthoc", "Perform Post-hoc Comparisons", TRUE),
                                          conditionalPanel(
                                            condition = "input.anova_posthoc",
                                            selectInput("anova_posthoc_method", "Post-hoc Method:",
                                                        choices = c("Tukey HSD" = "tukey",
                                                                    "Bonferroni" = "bonferroni",
                                                                    "Scheffe" = "scheffe",
                                                                    "Dunnett" = "dunnett"),
                                                        selected = "tukey")
                                          ),
                                          actionButton("run_anova", "Run ANOVA", 
                                                       class = "btn-primary btn-action",
                                                       icon = icon("play")),
                                          downloadButton("download_anova", "Download Results",
                                                         class = "btn-success btn-action")
                                   ),
                                   column(8, class = "main-panel",
                                          withSpinner(verbatimTextOutput("anova_results"), type = 4, color = "#0C5EA8"),
                                          withSpinner(DTOutput("anova_table"), type = 4, color = "#0C5EA8"),
                                          conditionalPanel(
                                            condition = "input.anova_assumptions",
                                            h4("ANOVA Assumptions Check"),
                                            withSpinner(verbatimTextOutput("anova_assumptions_results"), type = 4, color = "#0C5EA8"),
                                            withSpinner(plotOutput("anova_assumptions_plots"), type = 4, color = "#0C5EA8")
                                          ),
                                          conditionalPanel(
                                            condition = "input.anova_posthoc",
                                            h4("Post-hoc Comparisons"),
                                            withSpinner(verbatimTextOutput("anova_posthoc_results"), type = 4, color = "#0C5EA8"),
                                            withSpinner(DTOutput("anova_posthoc_table"), type = 4, color = "#0C5EA8"),
                                            withSpinner(plotOutput("anova_posthoc_plot"), type = 4, color = "#0C5EA8")
                                          )
                                   )
                                 )
                        ),
                        
                        tabPanel("Two-Way ANOVA",
                                 fluidRow(
                                   column(4, class = "sidebar-panel",
                                          selectInput("anova_outcome_2way", "Outcome Variable:", choices = NULL),
                                          selectInput("anova_factor1", "Factor 1:", choices = NULL),
                                          selectInput("anova_factor2", "Factor 2:", choices = NULL),
                                          checkboxInput("anova_interaction", "Include Interaction Term", TRUE),
                                          checkboxInput("anova_2way_posthoc", "Perform Post-hoc Tests", TRUE),
                                          actionButton("run_anova_2way", "Run Two-Way ANOVA", 
                                                       class = "btn-primary btn-action",
                                                       icon = icon("play")),
                                          downloadButton("download_anova_2way", "Download Results",
                                                         class = "btn-success btn-action")
                                   ),
                                   column(8, class = "main-panel",
                                          withSpinner(verbatimTextOutput("anova_2way_results"), type = 4, color = "#0C5EA8"),
                                          withSpinner(DTOutput("anova_2way_table"), type = 4, color = "#0C5EA8"),
                                          conditionalPanel(
                                            condition = "input.anova_2way_posthoc",
                                            h4("Simple Effects Analysis"),
                                            withSpinner(verbatimTextOutput("anova_2way_posthoc"), type = 4, color = "#0C5EA8")
                                          )
                                   )
                                 )
                        ),
                        
                        tabPanel("Welch ANOVA",
                                 fluidRow(
                                   column(4, class = "sidebar-panel",
                                          selectInput("welch_outcome", "Outcome Variable:", choices = NULL),
                                          selectInput("welch_group", "Grouping Variable:", choices = NULL),
                                          helpText("Welch ANOVA is robust to violations of homogeneity of variance."),
                                          actionButton("run_welch_anova", "Run Welch ANOVA", 
                                                       class = "btn-primary btn-action",
                                                       icon = icon("play")),
                                          downloadButton("download_welch_anova", "Download Results",
                                                         class = "btn-success btn-action")
                                   ),
                                   column(8, class = "main-panel",
                                          withSpinner(verbatimTextOutput("welch_anova_results"), type = 4, color = "#0C5EA8"),
                                          withSpinner(DTOutput("welch_anova_table"), type = 4, color = "#0C5EA8"),
                                          h4("Games-Howell Post-hoc Test"),
                                          withSpinner(verbatimTextOutput("welch_posthoc_results"), type = 4, color = "#0C5EA8")
                                   )
                                 )
                        ),
                        
                        tabPanel("ANOVA Assumptions",
                                 fluidRow(
                                   column(4, class = "sidebar-panel",
                                          selectInput("assumptions_outcome", "Outcome Variable:", choices = NULL),
                                          selectInput("assumptions_group", "Grouping Variable:", choices = NULL),
                                          h4("Tests for Assumptions"),
                                          checkboxInput("levene_test", "Levene's Test (Homogeneity of Variance)", TRUE),
                                          checkboxInput("normality_test", "Normality Tests", TRUE),
                                          checkboxInput("outlier_test", "Outlier Detection", TRUE),
                                          actionButton("run_assumptions", "Check Assumptions", 
                                                       class = "btn-primary btn-action",
                                                       icon = icon("check")),
                                          downloadButton("download_assumptions", "Download Report",
                                                         class = "btn-success btn-action")
                                   ),
                                   column(8, class = "main-panel",
                                          conditionalPanel(
                                            condition = "input.levene_test",
                                            h4("Homogeneity of Variance Tests"),
                                            withSpinner(verbatimTextOutput("levene_results"), type = 4, color = "#0C5EA8")
                                          ),
                                          conditionalPanel(
                                            condition = "input.normality_test",
                                            h4("Normality Tests"),
                                            withSpinner(verbatimTextOutput("normality_results"), type = 4, color = "#0C5EA8")
                                          ),
                                          conditionalPanel(
                                            condition = "input.outlier_test",
                                            h4("Outlier Detection"),
                                            withSpinner(verbatimTextOutput("outlier_results"), type = 4, color = "#0C5EA8")
                                          ),
                                          h4("Diagnostic Plots"),
                                          withSpinner(plotOutput("assumptions_plots"), type = 4, color = "#0C5EA8")
                                   )
                                 )
                        )
                      )
             ),
             
             # NEW: T-TESTS SUBTAB
             tabPanel("T-Tests",
                      tabsetPanel(
                        tabPanel("One-Sample T-Test",
                                 fluidRow(
                                   column(4, class = "sidebar-panel",
                                          selectInput("ttest_onesample_var", "Variable:", choices = NULL),
                                          numericInput("ttest_onesample_mu", "Test Value (μ₀):", value = 0),
                                          selectInput("ttest_onesample_alternative", "Alternative Hypothesis:",
                                                      choices = c("Two-sided" = "two.sided",
                                                                  "Greater than" = "greater",
                                                                  "Less than" = "less"),
                                                      selected = "two.sided"),
                                          sliderInput("ttest_onesample_conf", "Confidence Level:", 
                                                      min = 0.90, max = 0.99, value = 0.95, step = 0.01),
                                          actionButton("run_ttest_onesample", "Run One-Sample T-Test", 
                                                       class = "btn-primary btn-action",
                                                       icon = icon("play")),
                                          downloadButton("download_ttest_onesample", "Download Results",
                                                         class = "btn-success btn-action")
                                   ),
                                   column(8, class = "main-panel",
                                          withSpinner(verbatimTextOutput("ttest_onesample_results"), type = 4, color = "#0C5EA8"),
                                          withSpinner(DTOutput("ttest_onesample_table"), type = 4, color = "#0C5EA8"),
                                          h4("Descriptive Statistics"),
                                          withSpinner(verbatimTextOutput("ttest_onesample_descriptives"), type = 4, color = "#0C5EA8"),
                                          withSpinner(plotOutput("ttest_onesample_plot"), type = 4, color = "#0C5EA8")
                                   )
                                 )
                        ),
                        
                        tabPanel("Independent T-Test",
                                 fluidRow(
                                   column(4, class = "sidebar-panel",
                                          selectInput("ttest_independent_var", "Outcome Variable:", choices = NULL),
                                          selectInput("ttest_independent_group", "Grouping Variable:", choices = NULL),
                                          uiOutput("ttest_independent_levels_ui"),
                                          checkboxInput("ttest_var_equal", "Assume Equal Variances", FALSE),
                                          selectInput("ttest_alternative", "Alternative Hypothesis:",
                                                      choices = c("Two-sided" = "two.sided",
                                                                  "Greater than" = "greater",
                                                                  "Less than" = "less"),
                                                      selected = "two.sided"),
                                          sliderInput("ttest_conf", "Confidence Level:", 
                                                      min = 0.90, max = 0.99, value = 0.95, step = 0.01),
                                          actionButton("run_ttest_independent", "Run Independent T-Test", 
                                                       class = "btn-primary btn-action",
                                                       icon = icon("play")),
                                          downloadButton("download_ttest_independent", "Download Results",
                                                         class = "btn-success btn-action")
                                   ),
                                   column(8, class = "main-panel",
                                          withSpinner(verbatimTextOutput("ttest_independent_results"), type = 4, color = "#0C5EA8"),
                                          withSpinner(DTOutput("ttest_independent_table"), type = 4, color = "#0C5EA8"),
                                          h4("Group Descriptives"),
                                          withSpinner(verbatimTextOutput("ttest_independent_descriptives"), type = 4, color = "#0C5EA8"),
                                          h4("Variance Check"),
                                          withSpinner(verbatimTextOutput("ttest_variance_check"), type = 4, color = "#0C5EA8"),
                                          withSpinner(plotOutput("ttest_independent_plot"), type = 4, color = "#0C5EA8")
                                   )
                                 )
                        ),
                        
                        tabPanel("Paired T-Test",
                                 fluidRow(
                                   column(4, class = "sidebar-panel",
                                          selectInput("ttest_paired_var1", "Variable 1 (Before/Time 1):", choices = NULL),
                                          selectInput("ttest_paired_var2", "Variable 2 (After/Time 2):", choices = NULL),
                                          selectInput("ttest_paired_alternative", "Alternative Hypothesis:",
                                                      choices = c("Two-sided" = "two.sided",
                                                                  "Greater than" = "greater",
                                                                  "Less than" = "less"),
                                                      selected = "two.sided"),
                                          sliderInput("ttest_paired_conf", "Confidence Level:", 
                                                      min = 0.90, max = 0.99, value = 0.95, step = 0.01),
                                          actionButton("run_ttest_paired", "Run Paired T-Test", 
                                                       class = "btn-primary btn-action",
                                                       icon = icon("play")),
                                          downloadButton("download_ttest_paired", "Download Results",
                                                         class = "btn-success btn-action")
                                   ),
                                   column(8, class = "main-panel",
                                          withSpinner(verbatimTextOutput("ttest_paired_results"), type = 4, color = "#0C5EA8"),
                                          withSpinner(DTOutput("ttest_paired_table"), type = 4, color = "#0C5EA8"),
                                          h4("Paired Descriptives"),
                                          withSpinner(verbatimTextOutput("ttest_paired_descriptives"), type = 4, color = "#0C5EA8"),
                                          h4("Difference Analysis"),
                                          withSpinner(verbatimTextOutput("ttest_paired_differences"), type = 4, color = "#0C5EA8"),
                                          withSpinner(plotOutput("ttest_paired_plot"), type = 4, color = "#0C5EA8")
                                   )
                                 )
                        ),
                        
                        tabPanel("Non-parametric Tests",
                                 fluidRow(
                                   column(4, class = "sidebar-panel",
                                          h4("Mann-Whitney U Test"),
                                          selectInput("mannwhitney_var", "Outcome Variable:", choices = NULL),
                                          selectInput("mannwhitney_group", "Grouping Variable:", choices = NULL),
                                          actionButton("run_mannwhitney", "Run Mann-Whitney Test", 
                                                       class = "btn-primary btn-action",
                                                       icon = icon("play")),
                                          tags$hr(),
                                          h4("Wilcoxon Signed-Rank Test"),
                                          selectInput("wilcoxon_var1", "Variable 1:", choices = NULL),
                                          selectInput("wilcoxon_var2", "Variable 2:", choices = NULL),
                                          actionButton("run_wilcoxon", "Run Wilcoxon Test", 
                                                       class = "btn-primary btn-action",
                                                       icon = icon("play")),
                                          downloadButton("download_nonparametric", "Download Results",
                                                         class = "btn-success btn-action")
                                   ),
                                   column(8, class = "main-panel",
                                          h4("Mann-Whitney U Test Results"),
                                          withSpinner(verbatimTextOutput("mannwhitney_results"), type = 4, color = "#0C5EA8"),
                                          tags$hr(),
                                          h4("Wilcoxon Signed-Rank Test Results"),
                                          withSpinner(verbatimTextOutput("wilcoxon_results"), type = 4, color = "#0C5EA8")
                                   )
                                 )
                        )
                      )
             )
           )
  )
)

# Server function
server <- function(input, output, session) {
  
  # Add loading states for all action buttons
  observeEvent(input$load_data, {
    shinyjs::disable("load_data")
    shinyjs::html("load_data", "<i class='fas fa-spinner fa-spin'></i> Loading...")
  })
  
  observeEvent(input$run_freq, {
    shinyjs::disable("run_freq")
    shinyjs::html("run_freq", "<i class='fas fa-spinner fa-spin'></i> Analyzing...")
  })
  
  # Add to your button state management section (around line 600 in your original code)
  observeEvent(input$run_poisson, {
    shinyjs::disable("run_poisson")
    shinyjs::html("run_poisson", "<i class='fas fa-spinner fa-spin'></i> Running...")
  })
  
  observeEvent(poisson_results(), {
    shinyjs::enable("run_poisson")
    shinyjs::html("run_poisson", "<i class='fas fa-play'></i> Run Analysis")
  })
  
  observeEvent(input$run_central, {
    shinyjs::disable("run_central")
    shinyjs::html("run_central", "<i class='fas fa-spinner fa-spin'></i> Analyzing...")
  })
  
  observeEvent(input$run_epi, {
    shinyjs::disable("run_epi")
    shinyjs::html("run_epi", "<i class='fas fa-spinner fa-spin'></i> Generating...")
  })
  
  observeEvent(input$run_ip, {
    shinyjs::disable("run_ip")
    shinyjs::html("run_ip", "<i class='fas fa-spinner fa-spin'></i> Calculating...")
  })
  
  observeEvent(input$run_rr, {
    shinyjs::disable("run_rr")
    shinyjs::html("run_rr", "<i class='fas fa-spinner fa-spin'></i> Calculating...")
  })
  
  observeEvent(input$run_or, {
    shinyjs::disable("run_or")
    shinyjs::html("run_or", "<i class='fas fa-spinner fa-spin'></i> Calculating...")
  })
  
  observeEvent(input$run_surv, {
    shinyjs::disable("run_surv")
    shinyjs::html("run_surv", "<i class='fas fa-spinner fa-spin'></i> Analyzing...")
  })
  
  observeEvent(input$run_cox, {
    shinyjs::disable("run_cox")
    shinyjs::html("run_cox", "<i class='fas fa-spinner fa-spin'></i> Running...")
  })
  
  observeEvent(input$run_linear, {
    shinyjs::disable("run_linear")
    shinyjs::html("run_linear", "<i class='fas fa-spinner fa-spin'></i> Running...")
  })
  
  observeEvent(input$run_logistic, {
    shinyjs::disable("run_logistic")
    shinyjs::html("run_logistic", "<i class='fas fa-spinner fa-spin'></i> Running...")
  })
  
  observeEvent(input$run_chisq, {
    shinyjs::disable("run_chisq")
    shinyjs::html("run_chisq", "<i class='fas fa-spinner fa-spin'></i> Testing...")
  })
  
  # Reset button states after calculations
  observe({
    # Reset load data button
    if (!is.null(processed_data())) {
      shinyjs::enable("load_data")
      shinyjs::html("load_data", "<i class='fas fa-upload'></i> Load Data")
    }
  })
  
  # Reset analysis buttons after their respective calculations complete
  observeEvent(freq_results(), {
    shinyjs::enable("run_freq")
    shinyjs::html("run_freq", "<i class='fas fa-play'></i> Run Analysis")
  })
  
  observeEvent(central_results(), {
    shinyjs::enable("run_central")
    shinyjs::html("run_central", "<i class='fas fa-play'></i> Run Analysis")
  })
  
  observeEvent(epi_results(), {
    shinyjs::enable("run_epi")
    shinyjs::html("run_epi", "<i class='fas fa-play'></i> Run Analysis")
  })
  
  observeEvent(ip_results(), {
    shinyjs::enable("run_ip")
    shinyjs::html("run_ip", "<i class='fas fa-calculator'></i> Calculate")
  })
  
  observeEvent(rr_results(), {
    shinyjs::enable("run_rr")
    shinyjs::html("run_rr", "<i class='fas fa-calculator'></i> Calculate")
  })
  
  observeEvent(or_results(), {
    shinyjs::enable("run_or")
    shinyjs::html("run_or", "<i class='fas fa-calculator'></i> Calculate")
  })
  
  observeEvent(surv_results(), {
    shinyjs::enable("run_surv")
    shinyjs::html("run_surv", "<i class='fas fa-play'></i> Run Analysis")
  })
  
  observeEvent(cox_results(), {
    shinyjs::enable("run_cox")
    shinyjs::html("run_cox", "<i class='fas fa-play'></i> Run Analysis")
  })
  
  observeEvent(linear_results(), {
    shinyjs::enable("run_linear")
    shinyjs::html("run_linear", "<i class='fas fa-play'></i> Run Analysis")
  })
  
  observeEvent(logistic_results(), {
    shinyjs::enable("run_logistic")
    shinyjs::html("run_logistic", "<i class='fas fa-play'></i> Run Analysis")
  })
  
  observeEvent(chisq_results(), {
    shinyjs::enable("run_chisq")
    shinyjs::html("run_chisq", "<i class='fas fa-play'></i> Run Test")
  })
  
  # Reactive data with proper handling
  raw_data <- reactiveVal()
  processed_data <- reactiveVal()
  
  # Load data with proper error handling
  observeEvent(input$load_data, {
    req(input$file1)
    
    tryCatch({
      ext <- tools::file_ext(input$file1$name)
      
      df <- switch(ext,
                   csv = read.csv(input$file1$datapath,
                                  header = input$header,
                                  sep = input$sep,
                                  quote = input$quote,
                                  stringsAsFactors = FALSE),
                   xls = read_excel(input$file1$datapath, sheet = input$sheet),
                   xlsx = read_excel(input$file1$datapath, sheet = input$sheet),
                   sav = read_sav(input$file1$datapath),
                   dta = read_dta(input$file1$datapath),
                   sas7bdat = haven::read_sas(input$file1$datapath),
                   RData = {
                     env <- new.env()
                     load(input$file1$datapath, env)
                     get(ls(env)[1], env)
                   },
                   rda = {
                     env <- new.env()
                     load(input$file1$datapath, env)
                     get(ls(env)[1], env)
                   },
                   validate("Invalid file format")
      )
      
      # Convert to data.frame and handle labelled data
      if (inherits(df, "tbl_df")) {
        df <- as.data.frame(df)
      }
      
      # Handle labelled data from SPSS/Stata
      df <- convert_labelled(df)
      
      raw_data(df)
      
      # Process data - remove missing values
      df_processed <- df %>%
        mutate(across(everything(), ~ifelse(is.na(.), NA, .))) %>%
        na.omit()
      
      processed_data(df_processed)
      
      # Update all select inputs
      updateSelectInput(session, "freq_var", choices = names(df_processed))
      updateSelectInput(session, "central_var", choices = names(df_processed)[sapply(df_processed, is.numeric)])
      updateSelectInput(session, "epi_date", choices = names(df_processed))
      updateSelectInput(session, "epi_case", choices = names(df_processed))
      updateSelectInput(session, "epi_group", choices = c("None" = "none", names(df_processed)))
      updateSelectInput(session, "ip_case", choices = names(df_processed))
      updateSelectInput(session, "ip_pop", choices = names(df_processed)[sapply(df_processed, is.numeric)])
      updateSelectInput(session, "ip_time", choices = names(df_processed))
      updateSelectInput(session, "rr_outcome", choices = names(df_processed))
      updateSelectInput(session, "rr_exposure", choices = names(df_processed))
      updateSelectInput(session, "or_outcome", choices = names(df_processed))
      updateSelectInput(session, "or_exposure", choices = names(df_processed))
      updateSelectInput(session, "surv_time", choices = names(df_processed)[sapply(df_processed, is.numeric)])
      updateSelectInput(session, "surv_event", choices = names(df_processed))
      updateSelectInput(session, "surv_group", choices = c("None" = "none", names(df_processed)))
      updateSelectInput(session, "cox_time", choices = names(df_processed)[sapply(df_processed, is.numeric)])
      updateSelectInput(session, "cox_event", choices = names(df_processed))
      updateSelectInput(session, "cox_vars", choices = names(df_processed))
      updateSelectInput(session, "linear_outcome", choices = names(df_processed)[sapply(df_processed, is.numeric)])
      updateSelectInput(session, "linear_vars", choices = names(df_processed))
      updateSelectInput(session, "logistic_outcome", choices = names(df_processed))
      updateSelectInput(session, "logistic_vars", choices = names(df_processed))
      updateSelectInput(session, "chisq_var1", choices = names(df_processed))
      updateSelectInput(session, "chisq_var2", choices = names(df_processed))
      
      # Show success notification
      showNotification("Data loaded successfully!", type = "message", duration = 3)
      
    }, error = function(e) {
      showNotification(paste("Error loading file:", e$message), type = "error")
    })
  })
  
  # Missing values alert
  output$missing_values_alert <- renderUI({
    req(raw_data())
    df <- raw_data()
    missing_count <- sum(is.na(df))
    total_cells <- nrow(df) * ncol(df)
    missing_pct <- round((missing_count / total_cells) * 100, 1)
    
    if (missing_count > 0) {
      div(
        class = "alert alert-warning",
        h4("Missing Values Detected"),
        p(paste("Found", missing_count, "missing values (", missing_pct, "%) in the dataset.")),
        p("Missing values will be automatically excluded from all analyses.")
      )
    } else {
      div(
        class = "alert alert-success",
        h4("No Missing Values"),
        p("All data points are complete.")
      )
    }
  })
  
  # Data preview
  output$contents <- renderDT({
    req(processed_data())
    datatable(processed_data(), 
              options = list(scrollX = TRUE, pageLength = 10),
              rownames = FALSE)
  })
  
  # Data summary
  output$data_summary <- renderPrint({
    req(processed_data())
    df <- processed_data()
    cat("Dataset Summary:\n")
    cat("Number of observations:", nrow(df), "\n")
    cat("Number of variables:", ncol(df), "\n\n")
    cat("Variable types:\n")
    print(table(sapply(df, class)))
  })
  
  # Ultra-robust Frequency Analysis
  freq_results <- eventReactive(input$run_freq, {
    req(input$freq_var, processed_data())
    
    tryCatch({
      df <- processed_data()
      
      # Double-check variable exists
      if (!input$freq_var %in% names(df)) {
        return(list(
          table = data.frame(Error = "Variable not found in dataset"),
          success = FALSE
        ))
      }
      
      # Safely get the variable
      var_data <- df[[input$freq_var]]
      
      # Handle empty or all-NA data
      if (length(var_data) == 0 || all(is.na(var_data))) {
        return(list(
          table = data.frame(Error = "Variable contains no data"),
          success = FALSE
        ))
      }
      
      # Simple frequency table using base R (most reliable)
      freq_counts <- table(var_data, useNA = "ifany")
      freq_table <- data.frame(
        Category = names(freq_counts),
        Frequency = as.numeric(freq_counts),
        stringsAsFactors = FALSE
      )
      
      # Calculate proportions and percentages
      total <- sum(freq_table$Frequency)
      freq_table$Proportion <- freq_table$Frequency / total
      freq_table$Percentage <- freq_table$Proportion * 100
      
      # Clean up category names
      freq_table$Category[is.na(freq_table$Category)] <- "Missing"
      
      list(table = freq_table, success = TRUE)
      
    }, error = function(e) {
      return(list(
        table = data.frame(Error = paste("Error:", e$message)),
        success = FALSE
      ))
    })
  })
  
  output$freq_table <- renderDT({
    req(freq_results())
    
    results <- freq_results()
    
    if (!results$success) {
      # Show error in table format
      datatable(results$table, 
                options = list(dom = 't'),
                rownames = FALSE)
    } else {
      # Show frequency table
      datatable(results$table,
                options = list(
                  pageLength = 10,
                  dom = 'Blfrtip'
                ),
                rownames = FALSE) %>%
        formatRound(columns = c('Proportion', 'Percentage'), digits = 2)
    }
  })
  
  # ANOVA Table Output
  output$anova_table <- renderDT({
    req(anova_results())
    results <- anova_results()
    
    if (!results$success) {
      return(datatable(
        data.frame(Error = results$error),
        options = list(dom = 't'),
        rownames = FALSE
      ))
    }
    
    # Convert ANOVA summary to data frame
    anova_df <- as.data.frame(results$summary[[1]])
    anova_df <- cbind(Source = rownames(anova_df), anova_df)
    rownames(anova_df) <- NULL
    
    datatable(
      anova_df,
      options = list(
        dom = 't',
        pageLength = 10
      ),
      rownames = FALSE,
      caption = "One-Way ANOVA Results"
    ) %>%
      formatRound(columns = c('Df', 'Sum Sq', 'Mean Sq', 'F value', 'Pr(>F)'), digits = 4)
  })
  
  # ANOVA Post-hoc Table
  output$anova_posthoc_table <- renderDT({
    req(anova_results())
    results <- anova_results()
    
    if (!results$success || is.null(results$posthoc)) {
      return(datatable(
        data.frame(Message = "No post-hoc results available"),
        options = list(dom = 't'),
        rownames = FALSE
      ))
    }
    
    if (input$anova_posthoc_method == "tukey") {
      posthoc_df <- as.data.frame(results$posthoc$group_clean)
      posthoc_df <- cbind(Comparison = rownames(posthoc_df), posthoc_df)
      rownames(posthoc_df) <- NULL
      
      datatable(
        posthoc_df,
        options = list(
          dom = 't',
          pageLength = 10
        ),
        rownames = FALSE,
        caption = "Tukey HSD Post-hoc Comparisons"
      ) %>%
        formatRound(columns = c('diff', 'lwr', 'upr', 'p adj'), digits = 4)
    } else {
      # For other post-hoc methods
      datatable(
        as.data.frame(results$posthoc$p.value),
        options = list(
          dom = 't',
          pageLength = 10
        ),
        caption = "Post-hoc Comparisons"
      )
    }
  })
  
  # Two-Way ANOVA Table
  output$anova_2way_table <- renderDT({
    req(anova_2way_results())
    results <- anova_2way_results()
    
    if (!results$success) {
      return(datatable(
        data.frame(Error = results$error),
        options = list(dom = 't'),
        rownames = FALSE
      ))
    }
    
    anova_df <- as.data.frame(results$summary[[1]])
    anova_df <- cbind(Source = rownames(anova_df), anova_df)
    rownames(anova_df) <- NULL
    
    datatable(
      anova_df,
      options = list(
        dom = 't',
        pageLength = 10
      ),
      rownames = FALSE,
      caption = "Two-Way ANOVA Results"
    ) %>%
      formatRound(columns = c('Df', 'Sum Sq', 'Mean Sq', 'F value', 'Pr(>F)'), digits = 4)
  })
  
  # Welch ANOVA Table
  output$welch_anova_table <- renderDT({
    req(welch_anova_results())
    results <- welch_anova_results()
    
    if (!results$success) {
      return(datatable(
        data.frame(Error = results$error),
        options = list(dom = 't'),
        rownames = FALSE
      ))
    }
    
    welch_df <- data.frame(
      Statistic = c("F value", "Num DF", "Denom DF", "P-value"),
      Value = c(
        round(results$welch_test$statistic, 4),
        round(results$welch_test$parameter[1], 2),
        round(results$welch_test$parameter[2], 2),
        round(results$welch_test$p.value, 4)
      )
    )
    
    datatable(
      welch_df,
      options = list(dom = 't'),
      rownames = FALSE,
      caption = "Welch ANOVA Results"
    )
  })
  
  # One-Sample T-Test Table
  output$ttest_onesample_table <- renderDT({
    req(ttest_onesample_results())
    results <- ttest_onesample_results()
    
    if (!results$success) {
      return(datatable(
        data.frame(Error = results$error),
        options = list(dom = 't'),
        rownames = FALSE
      ))
    }
    
    ttest_df <- data.frame(
      Statistic = c("t value", "Degrees of Freedom", "P-value", "Confidence Level", "Alternative"),
      Value = c(
        round(results$ttest$statistic, 4),
        round(results$ttest$parameter, 2),
        round(results$ttest$p.value, 4),
        paste0(input$ttest_onesample_conf * 100, "%"),
        results$ttest$alternative
      )
    )
    
    datatable(
      ttest_df,
      options = list(dom = 't'),
      rownames = FALSE,
      caption = "One-Sample T-Test Results"
    )
  })
  
  # Independent T-Test Table
  output$ttest_independent_table <- renderDT({
    req(ttest_independent_results())
    results <- ttest_independent_results()
    
    if (!results$success) {
      return(datatable(
        data.frame(Error = results$error),
        options = list(dom = 't'),
        rownames = FALSE
      ))
    }
    
    ttest_df <- data.frame(
      Statistic = c("t value", "Degrees of Freedom", "P-value", "Confidence Level", "Alternative", "Equal Variances"),
      Value = c(
        round(results$ttest$statistic, 4),
        round(results$ttest$parameter, 2),
        round(results$ttest$p.value, 4),
        paste0(input$ttest_conf * 100, "%"),
        results$ttest$alternative,
        ifelse(input$ttest_var_equal, "Yes", "No")
      )
    )
    
    datatable(
      ttest_df,
      options = list(dom = 't'),
      rownames = FALSE,
      caption = "Independent T-Test Results"
    )
  })
  
  # Paired T-Test Table
  output$ttest_paired_table <- renderDT({
    req(ttest_paired_results())
    results <- ttest_paired_results()
    
    if (!results$success) {
      return(datatable(
        data.frame(Error = results$error),
        options = list(dom = 't'),
        rownames = FALSE
      ))
    }
    
    ttest_df <- data.frame(
      Statistic = c("t value", "Degrees of Freedom", "P-value", "Confidence Level", "Alternative"),
      Value = c(
        round(results$ttest$statistic, 4),
        round(results$ttest$parameter, 2),
        round(results$ttest$p.value, 4),
        paste0(input$ttest_paired_conf * 100, "%"),
        results$ttest$alternative
      )
    )
    
    datatable(
      ttest_df,
      options = list(dom = 't'),
      rownames = FALSE,
      caption = "Paired T-Test Results"
    )
  })
  
  output$freq_plot <- renderPlotly({
    req(freq_results())
    
    results <- freq_results()
    
    if (!results$success) {
      # Error plot
      p <- plot_ly() %>%
        add_annotations(
          text = results$table$Error[1],
          x = 0.5,
          y = 0.5,
          xref = "paper",
          yref = "paper",
          showarrow = FALSE,
          font = list(size = 16, color = "red")
        ) %>%
        layout(
          xaxis = list(showgrid = FALSE, zeroline = FALSE, showticklabels = FALSE),
          yaxis = list(showgrid = FALSE, zeroline = FALSE, showticklabels = FALSE)
        )
      return(p)
    }
    
    # Create bar plot using plot_ly directly (more reliable)
    p <- plot_ly(results$table,
                 x = ~Category,
                 y = ~Frequency,
                 type = 'bar',
                 marker = list(color = input$freq_color),
                 hoverinfo = 'text',
                 text = ~paste('Category:', Category,
                               '<br>Count:', Frequency,
                               '<br>Percentage:', round(Percentage, 1), '%')) %>%
      layout(
        title = list(text = input$freq_title, size = input$freq_title_size),
        xaxis = list(title = input$freq_var, tickangle = 45),
        yaxis = list(title = "Frequency"),
        margin = list(b = 100)
      )
    
    return(p)
  })
  
  # Add to the data loading observer (around line 1070 in your original code)
  observeEvent(processed_data(), {
    req(processed_data())
    df_processed <- processed_data()
    
    # Update Poisson regression inputs
    updateSelectInput(session, "poisson_outcome", choices = names(df_processed)[sapply(df_processed, is.numeric)])
    updateSelectInput(session, "poisson_vars", choices = names(df_processed))
    updateSelectInput(session, "poisson_offset_var", choices = names(df_processed)[sapply(df_processed, is.numeric)])
    
    # Update central tendency inputs
    updateSelectInput(session, "central_var", choices = names(df_processed)[sapply(df_processed, is.numeric)])
    updateSelectInput(session, "central_stratify", choices = c("None" = "none", names(df_processed)))
  })
  
  
  # Ultra-robust Central Tendency Analysis - UPDATED VERSION with stratification
  # Central tendency table output - ADD THIS SECTION
  output$central_table <- renderDT({
    req(central_results())
    
    results <- central_results()
    
    if (!results$success) {
      # Show error in table format
      datatable(results$table, 
                options = list(dom = 't'),
                rownames = FALSE,
                caption = "Central Tendency Analysis - Error")
    } else {
      # Show central tendency table
      datatable(results$table,
                options = list(
                  pageLength = 20,
                  dom = 'Blfrtip',
                  scrollX = TRUE
                ),
                rownames = FALSE,
                caption = "Central Tendency Measures") %>%
        formatRound(columns = c('Mean', 'Median', 'SD', 'Variance', 'Min', 'Max', 
                                'Q1', 'Q3', 'IQR', 'Skewness', 'Kurtosis'), 
                    digits = 4)
    }
  })
  central_results <- eventReactive(input$run_central, {
    req(input$central_var, processed_data())
    
    tryCatch({
      df <- processed_data()
      
      # Double-check variable exists
      if (!input$central_var %in% names(df)) {
        return(list(
          table = data.frame(Error = "Variable not found in dataset"),
          success = FALSE
        ))
      }
      
      # Safely get the variable
      var_data <- df[[input$central_var]]
      
      # Handle non-numeric data
      if (!is.numeric(var_data)) {
        return(list(
          table = data.frame(Error = "Selected variable is not numeric"),
          success = FALSE
        ))
      }
      
      # Handle empty or all-NA data
      if (length(var_data) == 0 || all(is.na(var_data))) {
        return(list(
          table = data.frame(Error = "Variable contains no data"),
          success = FALSE
        ))
      }
      
      # Handle stratification variable if selected
      stratify_data <- NULL
      group_levels <- "Overall"
      
      if (input$central_stratify != "none" && input$central_stratify %in% names(df)) {
        stratify_data <- df[[input$central_stratify]]
        
        # Convert to factor and handle NAs for stratification
        if (all(is.na(stratify_data))) {
          return(list(
            table = data.frame(Error = "Stratification variable contains no valid data"),
            success = FALSE
          ))
        }
        
        stratify_data <- as.factor(stratify_data)
        stratify_data[is.na(stratify_data)] <- "Missing"
        group_levels <- levels(stratify_data)
      }
      
      # Remove NA values for calculations
      if (!is.null(stratify_data)) {
        clean_data <- var_data[!is.na(var_data) & !is.na(stratify_data)]
        clean_stratify <- stratify_data[!is.na(var_data) & !is.na(stratify_data)]
      } else {
        clean_data <- var_data[!is.na(var_data)]
        clean_stratify <- NULL
      }
      
      if (length(clean_data) == 0) {
        return(list(
          table = data.frame(Error = "No valid numeric data after removing missing values"),
          success = FALSE
        ))
      }
      
      # Calculate central tendency measures with stratification
      get_mode <- function(x) {
        if (length(x) == 0) return(NA)
        ux <- unique(x)
        ux[which.max(tabulate(match(x, ux)))]
      }
      
      calculate_stats <- function(data, group_name = "Overall") {
        data.frame(
          Group = group_name,
          Mean = mean(data, na.rm = TRUE),
          Median = median(data, na.rm = TRUE),
          SD = sd(data, na.rm = TRUE),
          Variance = var(data, na.rm = TRUE),
          Min = min(data, na.rm = TRUE),
          Max = max(data, na.rm = TRUE),
          Q1 = quantile(data, 0.25, na.rm = TRUE),
          Q3 = quantile(data, 0.75, na.rm = TRUE),
          IQR = IQR(data, na.rm = TRUE),
          Skewness = ifelse(length(data) > 2, moments::skewness(data, na.rm = TRUE), NA),
          Kurtosis = ifelse(length(data) > 3, moments::kurtosis(data, na.rm = TRUE), NA),
          Mode = get_mode(data),
          Count = length(data),
          Missing = sum(is.na(data)),
          stringsAsFactors = FALSE
        )
      }
      
      if (!is.null(clean_stratify)) {
        # Calculate statistics for each group
        central_table <- do.call(rbind, lapply(levels(clean_stratify), function(level) {
          group_data <- clean_data[clean_stratify == level]
          if (length(group_data) > 0) {
            calculate_stats(group_data, level)
          }
        }))
        
        # Add overall statistics
        overall_stats <- calculate_stats(clean_data, "Overall")
        central_table <- rbind(overall_stats, central_table)
        
      } else {
        # Calculate overall statistics only
        central_table <- calculate_stats(clean_data, "Overall")
      }
      
      # Format numeric values
      numeric_cols <- c("Mean", "Median", "SD", "Variance", "Min", "Max", "Q1", "Q3", "IQR", "Skewness", "Kurtosis")
      central_table[numeric_cols] <- lapply(central_table[numeric_cols], function(x) round(x, 4))
      
      list(
        table = central_table, 
        data = clean_data, 
        stratify = clean_stratify,
        group_levels = group_levels,
        success = TRUE
      )
      
    }, error = function(e) {
      return(list(
        table = data.frame(Error = paste("Error:", e$message)),
        success = FALSE
      ))
    })
  })
  
  # Updated boxplot output with stratification
  output$central_boxplot <- renderPlotly({
    req(central_results())
    
    results <- central_results()
    
    if (!results$success) {
      # Error plot
      p <- plot_ly() %>%
        add_annotations(
          text = results$table$Error[1],
          x = 0.5,
          y = 0.5,
          xref = "paper",
          yref = "paper",
          showarrow = FALSE,
          font = list(size = 16, color = "red")
        ) %>%
        layout(
          title = "Boxplot - Error",
          xaxis = list(showgrid = FALSE, zeroline = FALSE, showticklabels = FALSE),
          yaxis = list(showgrid = FALSE, zeroline = FALSE, showticklabels = FALSE)
        )
      return(p)
    }
    
    if (!is.null(results$stratify)) {
      # Create grouped box plot
      plot_data <- data.frame(
        value = results$data,
        group = results$stratify
      )
      
      p <- plot_ly(plot_data, 
                   x = ~group, 
                   y = ~value, 
                   type = "box",
                   color = ~group,
                   boxpoints = "all",
                   jitter = 0.3,
                   pointpos = -1.8,
                   marker = list(size = 3)) %>%
        layout(
          title = list(text = paste(input$central_title, "- Boxplot by", input$central_stratify), 
                       size = input$central_title_size),
          yaxis = list(title = input$central_var),
          xaxis = list(title = input$central_stratify),
          showlegend = FALSE
        )
    } else {
      # Create single box plot
      p <- plot_ly(y = ~results$data,
                   type = "box",
                   name = input$central_var,
                   boxpoints = "all",
                   jitter = 0.3,
                   pointpos = -1.8,
                   marker = list(color = input$central_color, size = 3),
                   line = list(color = input$central_color),
                   fillcolor = paste0(input$central_color, "33")) %>%
        layout(
          title = list(text = paste(input$central_title, "- Boxplot"), 
                       size = input$central_title_size),
          yaxis = list(title = input$central_var),
          xaxis = list(title = "", showticklabels = FALSE),
          showlegend = FALSE
        )
    }
    
    return(p)
  })
  
  # Updated histogram output with stratification
  output$central_histogram <- renderPlotly({
    req(central_results())
    
    results <- central_results()
    
    if (!results$success) {
      # Error plot
      p <- plot_ly() %>%
        add_annotations(
          text = results$table$Error[1],
          x = 0.5,
          y = 0.5,
          xref = "paper",
          yref = "paper",
          showarrow = FALSE,
          font = list(size = 16, color = "red")
        ) %>%
        layout(
          title = "Histogram - Error",
          xaxis = list(showgrid = FALSE, zeroline = FALSE, showticklabels = FALSE),
          yaxis = list(showgrid = FALSE, zeroline = FALSE, showticklabels = FALSE)
        )
      return(p)
    }
    
    if (!is.null(results$stratify)) {
      # Create grouped histogram
      plot_data <- data.frame(
        value = results$data,
        group = results$stratify
      )
      
      p <- plot_ly(plot_data) %>%
        add_histogram(x = ~value, color = ~group, opacity = 0.7) %>%
        layout(
          title = list(text = paste(input$central_title, "- Histogram by", input$central_stratify), 
                       size = input$central_title_size),
          xaxis = list(title = input$central_var),
          yaxis = list(title = "Frequency"),
          barmode = "overlay",
          bargap = 0.1
        )
    } else {
      # Create single histogram
      p <- plot_ly(x = ~results$data,
                   type = "histogram",
                   nbinsx = input$central_bins,
                   marker = list(color = input$central_color,
                                 line = list(color = "white", width = 1)),
                   opacity = 0.7) %>%
        layout(
          title = list(text = paste(input$central_title, "- Histogram"), 
                       size = input$central_title_size),
          xaxis = list(title = input$central_var),
          yaxis = list(title = "Frequency"),
          bargap = 0.1
        )
    }
    
    return(p)
  })
  
  # CORRECTED DOWNLOAD HANDLERS FOR DESCRIPTIVE ANALYSIS
  
  # Frequency Analysis Download
  output$download_freq <- downloadHandler(
    filename = function() {
      if(input$freq_type == "Table") {
        paste("frequency_analysis_", input$freq_var, "_", Sys.Date(), ".docx", sep = "")
      } else {
        paste("frequency_plot_", input$freq_var, "_", Sys.Date(), ".png", sep = "")
      }
    },
    content = function(file) {
      req(freq_results())
      
      results <- freq_results()
      
      if(input$freq_type == "Table") {
        if (!results$success) {
          # Create error document
          ft <- flextable(data.frame(Error = "Unable to generate frequency table due to error")) %>%
            set_caption("Frequency Analysis - Error") %>%
            theme_zebra()
          
          doc <- read_docx() %>% 
            body_add_flextable(ft)
        } else {
          # Create formatted table for download
          ft <- flextable(results$table) %>%
            set_caption(paste("Frequency Distribution -", input$freq_var)) %>%
            theme_zebra() %>%
            autofit()
          
          doc <- read_docx() %>% 
            body_add_flextable(ft)
        }
        print(doc, target = file)
        
      } else {
        # Download plot
        if (!results$success) {
          # Create error plot
          p <- ggplot() +
            annotate("text", x = 1, y = 1, 
                     label = "Error: Unable to generate frequency plot", 
                     size = 6, color = "red") +
            theme_void()
        } else {
          # Create high-quality bar plot using ggplot2
          p <- ggplot(results$table, aes(x = Category, y = Frequency)) +
            geom_bar(stat = "identity", fill = input$freq_color, alpha = 0.8) +
            labs(title = input$freq_title,
                 x = input$freq_var,
                 y = "Frequency") +
            theme_minimal() +
            theme(
              plot.title = element_text(size = input$freq_title_size, face = "bold"),
              axis.title.x = element_text(size = input$freq_x_size),
              axis.title.y = element_text(size = input$freq_y_size),
              axis.text.x = element_text(angle = 45, hjust = 1)
            )
        }
        ggsave(file, plot = p, device = "png", width = 10, height = 6, dpi = 300)
      }
    }
  )
  
  # Central Tendency Download (already exists in your code, but ensure it's properly placed)
  output$download_central <- downloadHandler(
    filename = function() {
      if(input$central_type == "Table") {
        paste("central_tendency_", input$central_var, "_", Sys.Date(), ".docx", sep = "")
      } else if(input$central_type == "Boxplot") {
        paste("boxplot_", input$central_var, "_", Sys.Date(), ".png", sep = "")
      } else {
        paste("histogram_", input$central_var, "_", Sys.Date(), ".png", sep = "")
      }
    },
    content = function(file) {
      req(central_results())
      
      results <- central_results()
      
      if(input$central_type == "Table") {
        if (!results$success) {
          # Create error document
          ft <- flextable(data.frame(Error = "Unable to generate central tendency table due to error")) %>%
            set_caption("Central Tendency Analysis - Error") %>%
            theme_zebra()
          
          doc <- read_docx() %>% 
            body_add_flextable(ft)
        } else {
          # Create formatted table for download
          ft <- flextable(results$table) %>%
            set_caption(paste("Central Tendency Measures -", input$central_var)) %>%
            theme_zebra() %>%
            autofit()
          
          doc <- read_docx() %>% 
            body_add_flextable(ft)
        }
        print(doc, target = file)
        
      } else if(input$central_type == "Boxplot") {
        # Download boxplot
        if (!results$success) {
          # Create error plot
          p <- ggplot() +
            annotate("text", x = 1, y = 1, 
                     label = "Error: Unable to generate boxplot", 
                     size = 6, color = "red") +
            theme_void()
        } else {
          # Create boxplot using ggplot2 for high-quality download
          df_plot <- data.frame(value = results$data)
          if (!is.null(results$stratify)) {
            df_plot$group <- results$stratify
            p <- ggplot(df_plot, aes(x = group, y = value, fill = group)) +
              geom_boxplot(alpha = 0.7, outlier.color = "red") +
              labs(title = paste(input$central_title, "- Boxplot by", input$central_stratify),
                   y = input$central_var,
                   x = input$central_stratify) +
              theme_minimal() +
              theme(
                plot.title = element_text(size = input$central_title_size, face = "bold"),
                legend.position = "none"
              )
          } else {
            p <- ggplot(df_plot, aes(y = value)) +
              geom_boxplot(fill = input$central_color, alpha = 0.7, outlier.color = "red") +
              labs(title = paste(input$central_title, "- Boxplot"),
                   y = input$central_var) +
              theme_minimal() +
              theme(
                plot.title = element_text(size = input$central_title_size, face = "bold"),
                axis.title.x = element_blank(),
                axis.text.x = element_blank(),
                axis.ticks.x = element_blank()
              )
          }
        }
        ggsave(file, plot = p, device = "png", width = 8, height = 6, dpi = 300)
        
      } else {
        # Download histogram
        if (!results$success) {
          # Create error plot
          p <- ggplot() +
            annotate("text", x = 1, y = 1, 
                     label = "Error: Unable to generate histogram", 
                     size = 6, color = "red") +
            theme_void()
        } else {
          # Create histogram using ggplot2 for high-quality download
          df_plot <- data.frame(value = results$data)
          if (!is.null(results$stratify)) {
            df_plot$group <- results$stratify
            p <- ggplot(df_plot, aes(x = value, fill = group)) +
              geom_histogram(alpha = 0.7, position = "identity", bins = input$central_bins) +
              labs(title = paste(input$central_title, "- Histogram by", input$central_stratify),
                   x = input$central_var, 
                   y = "Frequency") +
              theme_minimal() +
              theme(
                plot.title = element_text(size = input$central_title_size, face = "bold")
              )
          } else {
            p <- ggplot(df_plot, aes(x = value)) +
              geom_histogram(fill = input$central_color, bins = input$central_bins, 
                             alpha = 0.7, color = "white") +
              labs(title = paste(input$central_title, "- Histogram"),
                   x = input$central_var, 
                   y = "Frequency") +
              theme_minimal() +
              theme(
                plot.title = element_text(size = input$central_title_size, face = "bold")
              )
          }
        }
        ggsave(file, plot = p, device = "png", width = 10, height = 6, dpi = 300)
      }
    }
  )
  
  # Epidemic Curve Plot Download
  output$download_epicurve <- downloadHandler(
    filename = function() {
      paste("epidemic_curve_", input$epi_interval, "_", Sys.Date(), ".png", sep = "")
    },
    content = function(file) {
      req(epi_results())
      
      results <- epi_results()
      
      if (!results$success) {
        # Create error plot
        p <- ggplot() +
          annotate("text", x = 1, y = 1, 
                   label = paste("Error:", results$error), 
                   size = 5, color = "red") +
          theme_void()
      } else {
        epi_data <- results$data
        
        # Create high-quality epidemic curve using ggplot2
        if (input$epi_group != "none") {
          p <- ggplot(epi_data, aes(x = time_group, y = cases, fill = Group)) +
            geom_col(position = "stack") +
            scale_fill_brewer(palette = "Set1") +
            labs(
              title = input$epi_title,
              subtitle = paste("Time Interval:", input$epi_interval),
              x = "Date",
              y = "Number of Cases",
              fill = input$epi_group
            )
        } else {
          p <- ggplot(epi_data, aes(x = time_group, y = cases)) +
            geom_col(fill = input$epi_color, alpha = 0.8) +
            labs(
              title = input$epi_title,
              subtitle = paste("Time Interval:", input$epi_interval),
              x = "Date",
              y = "Number of Cases"
            )
        }
        
        p <- p +
          theme_minimal() +
          theme(
            plot.title = element_text(size = input$epi_title_size, face = "bold"),
            plot.subtitle = element_text(size = input$epi_title_size - 2, color = "gray50"),
            axis.title.x = element_text(size = input$epi_x_size),
            axis.title.y = element_text(size = input$epi_y_size),
            axis.text.x = element_text(angle = 45, hjust = 1),
            legend.position = ifelse(input$epi_group != "none", "right", "none")
          )
      }
      
      ggsave(file, plot = p, device = "png", width = 12, height = 8, dpi = 300)
    }
  )
  
  # Epidemic Curve Data Download
  output$download_epi_data <- downloadHandler(
    filename = function() {
      paste("epidemic_data_", input$epi_interval, "_", Sys.Date(), ".csv", sep = "")
    },
    content = function(file) {
      req(epi_results())
      
      results <- epi_results()
      
      if (!results$success) {
        # Create error data frame
        write.csv(data.frame(Error = results$error), file, row.names = FALSE)
      } else {
        # Write epidemic data to CSV
        write.csv(results$data, file, row.names = FALSE)
      }
    }
  )
  # Ultra-robust Epidemic Curve Analysis - CORRECTED VERSION
  epi_results <- eventReactive(input$run_epi, {
    req(input$epi_date, input$epi_case, processed_data())
    
    tryCatch({
      df <- processed_data()
      
      # Double-check variables exist
      if (!input$epi_date %in% names(df)) {
        return(list(
          data = data.frame(),
          error = "Date variable not found in dataset",
          success = FALSE
        ))
      }
      
      if (!input$epi_case %in% names(df)) {
        return(list(
          data = data.frame(),
          error = "Case count variable not found in dataset",
          success = FALSE
        ))
      }
      
      if (input$epi_group != "none" && !input$epi_group %in% names(df)) {
        return(list(
          data = data.frame(),
          error = "Grouping variable not found in dataset",
          success = FALSE
        ))
      }
      
      # Safely get the variables
      date_data <- df[[input$epi_date]]
      case_data <- df[[input$epi_case]]
      
      # Handle empty or all-NA data
      if (length(date_data) == 0 || all(is.na(date_data))) {
        return(list(
          data = data.frame(),
          error = "Date variable contains no data",
          success = FALSE
        ))
      }
      
      if (length(case_data) == 0 || all(is.na(case_data))) {
        return(list(
          data = data.frame(),
          error = "Case count variable contains no data",
          success = FALSE
        ))
      }
      
      # NEW APPROACH: Handle date variable as categorical for all interval types
      # Convert date variable to character/factor for categorical handling
      df$date_converted <- as.character(date_data)
      
      # Handle missing values in date
      df$date_converted[is.na(df$date_converted)] <- "Missing"
      
      # Ensure case count is numeric
      df$case_numeric <- as.numeric(case_data)
      
      if (all(is.na(df$case_numeric))) {
        return(list(
          data = data.frame(),
          error = "Case count variable could not be converted to numeric",
          success = FALSE
        ))
      }
      
      # Handle grouping variable
      if (input$epi_group != "none") {
        group_data <- df[[input$epi_group]]
        if (all(is.na(group_data))) {
          return(list(
            data = data.frame(),
            error = "Grouping variable contains no valid data",
            success = FALSE
          ))
        }
        df$group_var <- as.character(group_data)
        df$group_var[is.na(df$group_var)] <- "Missing"
      }
      
      # NEW: For Day/Week/Month/Year intervals, treat date as categorical
      # The aggregation will be done by the unique date values as categories
      
      # Remove rows with invalid case counts
      df_clean <- df %>%
        filter(!is.na(case_numeric))
      
      if (nrow(df_clean) == 0) {
        return(list(
          data = data.frame(),
          error = "No valid data after removing missing case counts",
          success = FALSE
        ))
      }
      
      # Use date_converted directly as time_group for categorical aggregation
      df_clean$time_group <- df_clean$date_converted
      
      # For date-like categorical variables, we can try to convert to proper dates
      # but if it fails, we'll keep them as categorical
      if (input$epi_interval %in% c("Day", "Week", "Month", "Year")) {
        # Try to parse as dates, but if it fails, keep as categorical
        tryCatch({
          # Sample a few values to check if they look like dates
          sample_dates <- head(na.omit(unique(df_clean$date_converted)), 10)
          
          # Check if any sample looks like a date
          looks_like_date <- any(
            grepl("\\d{1,4}[/-]\\d{1,2}[/-]\\d{1,4}", sample_dates) |
              grepl("\\d{1,2}[/-]\\d{1,2}[/-]\\d{2,4}", sample_dates) |
              grepl("jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec", tolower(sample_dates), ignore.case = TRUE)
          )
          
          if (looks_like_date) {
            # Try to convert to dates for proper interval grouping
            convert_date <- function(x) {
              # Try different date formats
              formats <- c("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%Y.%m.%d", 
                           "%d-%m-%Y", "%m-%d-%Y", "%B %d, %Y", "%d %B %Y",
                           "%b %d, %Y", "%d %b %Y", "%Y/%m/%d", "%m/%d/%y")
              
              for (fmt in formats) {
                parsed <- as.Date(x, format = fmt)
                if (!all(is.na(parsed))) {
                  return(parsed)
                }
              }
              return(as.Date(NA))
            }
            
            df_clean$date_attempt <- convert_date(df_clean$date_converted)
            
            # Only use date conversion if it worked for at least some values
            if (!all(is.na(df_clean$date_attempt))) {
              if (input$epi_interval == "Day") {
                df_clean$time_group <- df_clean$date_attempt
              } else if (input$epi_interval == "Week") {
                df_clean$time_group <- lubridate::floor_date(df_clean$date_attempt, "week")
              } else if (input$epi_interval == "Month") {
                df_clean$time_group <- lubridate::floor_date(df_clean$date_attempt, "month")
              } else {
                df_clean$time_group <- lubridate::floor_date(df_clean$date_attempt, "year")
              }
              
              # Convert back to character for consistent handling
              df_clean$time_group <- as.character(df_clean$time_group)
            }
          }
        }, error = function(e) {
          # If date conversion fails, keep as categorical
          message("Date conversion failed, using categorical approach: ", e$message)
        })
      }
      
      # Aggregate data
      if (input$epi_group != "none") {
        epi_data <- df_clean %>%
          group_by(time_group, group_var) %>%
          summarise(
            cases = sum(case_numeric, na.rm = TRUE),
            n_records = n(),
            .groups = "drop"
          ) %>%
          rename(Group = group_var)
      } else {
        epi_data <- df_clean %>%
          group_by(time_group) %>%
          summarise(
            cases = sum(case_numeric, na.rm = TRUE),
            n_records = n(),
            .groups = "drop"
          )
      }
      
      # Remove any time groups with no cases
      epi_data <- epi_data %>%
        filter(cases > 0)
      
      if (nrow(epi_data) == 0) {
        return(list(
          data = data.frame(),
          error = "No cases found in the specified date range",
          success = FALSE
        ))
      }
      
      # Sort by time_group for proper ordering
      epi_data <- epi_data %>%
        arrange(time_group)
      
      # Calculate summary statistics
      summary_stats <- list(
        total_cases = sum(epi_data$cases),
        time_range = paste(min(epi_data$time_group), "to", max(epi_data$time_group)),
        n_time_periods = nrow(epi_data),
        peak_cases = max(epi_data$cases),
        peak_time = epi_data$time_group[which.max(epi_data$cases)]
      )
      
      list(
        data = epi_data,
        summary = summary_stats,
        success = TRUE
      )
      
    }, error = function(e) {
      return(list(
        data = data.frame(),
        error = paste("Error:", e$message),
        success = FALSE
      ))
    })
  })
  
  output$epi_curve <- renderPlotly({
    req(epi_results())
    
    results <- epi_results()
    
    if (!results$success) {
      # Error plot
      p <- plot_ly() %>%
        add_annotations(
          text = results$error,
          x = 0.5,
          y = 0.5,
          xref = "paper",
          yref = "paper",
          showarrow = FALSE,
          font = list(size = 14, color = "red")
        ) %>%
        layout(
          title = "Epidemic Curve - Error",
          xaxis = list(showgrid = FALSE, zeroline = FALSE, showticklabels = FALSE),
          yaxis = list(showgrid = FALSE, zeroline = FALSE, showticklabels = FALSE)
        )
      return(p)
    }
    
    epi_data <- results$data
    
    # Create epidemic curve
    if (input$epi_group != "none") {
      # Grouped epidemic curve
      p <- plot_ly(epi_data, 
                   x = ~time_group, 
                   y = ~cases, 
                   color = ~Group,
                   type = 'bar',
                   hoverinfo = 'text',
                   text = ~paste('Date:', time_group,
                                 '<br>Group:', Group,
                                 '<br>Cases:', cases,
                                 '<br>Records:', n_records)) %>%
        layout(
          title = list(text = input$epi_title, size = input$epi_title_size),
          xaxis = list(title = "Date", tickangle = 45),
          yaxis = list(title = "Number of Cases"),
          barmode = 'stack',
          margin = list(b = 100, l = 60, r = 40, t = 60),
          showlegend = TRUE
        )
    } else {
      # Simple epidemic curve
      p <- plot_ly(epi_data, 
                   x = ~time_group, 
                   y = ~cases,
                   type = 'bar',
                   marker = list(color = input$epi_color),
                   hoverinfo = 'text',
                   text = ~paste('Date:', time_group,
                                 '<br>Cases:', cases,
                                 '<br>Records:', n_records)) %>%
        layout(
          title = list(text = input$epi_title, size = input$epi_title_size),
          xaxis = list(title = "Date", tickangle = 45),
          yaxis = list(title = "Number of Cases"),
          margin = list(b = 100, l = 60, r = 40, t = 60),
          showlegend = FALSE
        )
    }
    
    return(p)
  })
  
  # Add summary statistics output
  output$epi_summary <- renderUI({
    req(epi_results())
    
    results <- epi_results()
    
    if (!results$success) {
      return(
        div(
          class = "alert alert-danger",
          h4("Analysis Error"),
          p(results$error)
        )
      )
    }
    
    summary <- results$summary
    
    div(
      class = "panel panel-default",
      div(
        class = "panel-heading",
        h4("Epidemic Summary Statistics", class = "panel-title")
      ),
      div(
        class = "panel-body",
        fluidRow(
          column(6,
                 tags$ul(
                   tags$li(tags$strong("Total Cases:"), summary$total_cases),
                   tags$li(tags$strong("Date Range:"), summary$date_range),
                   tags$li(tags$strong("Time Periods:"), summary$n_time_periods)
                 )
          ),
          column(6,
                 tags$ul(
                   tags$li(tags$strong("Peak Cases:"), summary$peak_cases),
                   tags$li(tags$strong("Peak Date:"), summary$peak_date)
                 )
          )
        )
      )
    )
  })
  
  output$download_epicurve <- downloadHandler(
    filename = function() {
      paste("epidemic_curve_", input$epi_interval, "_", Sys.Date(), ".png", sep = "")
    },
    content = function(file) {
      req(epi_results())
      
      results <- epi_results()
      
      if (!results$success) {
        # Create error plot
        p <- ggplot() +
          annotate("text", x = 1, y = 1, 
                   label = paste("Error:", results$error), 
                   size = 5, color = "red") +
          theme_void()
      } else {
        epi_data <- results$data
        
        # Create high-quality epidemic curve using ggplot2
        if (input$epi_group != "none") {
          p <- ggplot(epi_data, aes(x = time_group, y = cases, fill = Group)) +
            geom_col(position = "stack") +
            scale_fill_brewer(palette = "Set1") +
            labs(
              title = input$epi_title,
              subtitle = paste("Time Interval:", input$epi_interval),
              x = "Date",
              y = "Number of Cases",
              fill = input$epi_group
            )
        } else {
          p <- ggplot(epi_data, aes(x = time_group, y = cases)) +
            geom_col(fill = input$epi_color, alpha = 0.8) +
            labs(
              title = input$epi_title,
              subtitle = paste("Time Interval:", input$epi_interval),
              x = "Date",
              y = "Number of Cases"
            )
        }
        
        p <- p +
          theme_minimal() +
          theme(
            plot.title = element_text(size = input$epi_title_size, face = "bold"),
            plot.subtitle = element_text(size = input$epi_title_size - 2, color = "gray50"),
            axis.title.x = element_text(size = input$epi_x_size),
            axis.title.y = element_text(size = input$epi_y_size),
            axis.text.x = element_text(angle = 45, hjust = 1),
            legend.position = ifelse(input$epi_group != "none", "right", "none")
          )
      }
      
      ggsave(file, plot = p, device = "png", width = 12, height = 8, dpi = 300)
    }
  )
  
  # Also add data table download
  output$download_epi_data <- downloadHandler(
    filename = function() {
      paste("epidemic_data_", input$epi_interval, "_", Sys.Date(), ".csv", sep = "")
    },
    content = function(file) {
      req(epi_results())
      
      results <- epi_results()
      
      if (!results$success) {
        # Create error data frame
        write.csv(data.frame(Error = results$error), file, row.names = FALSE)
      } else {
        # Write epidemic data to CSV
        write.csv(results$data, file, row.names = FALSE)
      }
    }
  )
  
  # Incidence/Prevalence with action button
  ip_results <- eventReactive(input$run_ip, {
    req(input$ip_case, input$ip_pop, processed_data())
    df <- processed_data()
    
    if(input$ip_measure == "Incidence Rate") {
      req(input$ip_time)
      total_cases <- sum(df[[input$ip_case]], na.rm = TRUE)
      total_pop <- sum(df[[input$ip_pop]], na.rm = TRUE)
      total_time <- sum(df[[input$ip_time]], na.rm = TRUE)
      
      rate <- (total_cases / total_time) * 1000  # per 1000 person-time
      
      result <- data.frame(
        Measure = "Incidence Rate",
        Cases = total_cases,
        Population = total_pop,
        PersonTime = total_time,
        Rate = round(rate, 2),
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
        Rate = round(ci, 2),
        Units = "per 1000 population"
      )
    } else {
      total_cases <- sum(df[[input$ip_case]], na.rm = TRUE)
      total_pop <- sum(df[[input$ip_pop]], na.rm = TRUE)
      
      prevalence <- (total_cases / total_pop) * 100  # percentage
      
      result <- data.frame(
        Measure = "Prevalence",
        Cases = total_cases,
        Population = total_pop,
        Rate = round(prevalence, 2),
        Units = "percent"
      )
    }
    
    list(table = result)
  })
  
  output$ip_table <- renderDT({
    req(ip_results())
    datatable(ip_results()$table, options = list(dom = 't'))
  })
  
  output$download_ip <- downloadHandler(
    filename = "incidence_prevalence.docx",
    content = function(file) {
      ft <- flextable(ip_results()$table) %>%
        set_caption("Disease Frequency Measures") %>%
        theme_zebra()
      
      doc <- read_docx() %>% 
        body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # 4.0 COMPLETE DOWNLOAD HANDLERS
  
  # One-Way ANOVA Download
  output$download_anova <- downloadHandler(
    filename = function() {
      paste("oneway_anova_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(anova_results())
      results <- anova_results()
      
      doc <- read_docx()
      
      if (!results$success) {
        doc <- doc %>%
          body_add_par("ONE-WAY ANOVA ANALYSIS - ERROR", style = "heading 1") %>%
          body_add_par(results$error, style = "Normal")
      } else {
        # Title and overview
        doc <- doc %>%
          body_add_par("ONE-WAY ANOVA ANALYSIS REPORT", style = "heading 1") %>%
          body_add_par(paste("Outcome Variable:", input$anova_outcome), style = "Normal") %>%
          body_add_par(paste("Grouping Variable:", input$anova_group), style = "Normal") %>%
          body_add_par(paste("Number of Groups:", length(unique(processed_data()[[input$anova_group]]))), style = "Normal") %>%
          body_add_par("", style = "Normal")
        
        # ANOVA results table
        doc <- doc %>%
          body_add_par("ANOVA RESULTS", style = "heading 2")
        
        anova_df <- as.data.frame(results$summary[[1]])
        anova_df <- cbind(Source = rownames(anova_df), anova_df)
        rownames(anova_df) <- NULL
        
        ft_anova <- flextable(anova_df) %>%
          theme_zebra() %>%
          set_caption("One-Way ANOVA Table") %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_anova)
        
        # Group descriptives
        doc <- doc %>%
          body_add_par("GROUP DESCRIPTIVES", style = "heading 2")
        
        df <- processed_data()
        outcome <- df[[input$anova_outcome]]
        group <- df[[input$anova_group]]
        
        descriptives <- df %>%
          group_by(Group = !!sym(input$anova_group)) %>%
          summarise(
            N = sum(!is.na(!!sym(input$anova_outcome))),
            Mean = round(mean(!!sym(input$anova_outcome), na.rm = TRUE), 3),
            SD = round(sd(!!sym(input$anova_outcome), na.rm = TRUE), 3),
            SE = round(sd(!!sym(input$anova_outcome), na.rm = TRUE) / sqrt(sum(!is.na(!!sym(input$anova_outcome)))), 3),
            .groups = "drop"
          )
        
        ft_desc <- flextable(descriptives) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_desc)
        
        # Post-hoc results if available
        if (!is.null(results$posthoc)) {
          doc <- doc %>%
            body_add_par("POST-HOC COMPARISONS", style = "heading 2")
          
          if (input$anova_posthoc_method == "tukey") {
            posthoc_df <- as.data.frame(results$posthoc$group_clean)
            posthoc_df <- cbind(Comparison = rownames(posthoc_df), posthoc_df)
            rownames(posthoc_df) <- NULL
            
            ft_posthoc <- flextable(posthoc_df) %>%
              theme_zebra() %>%
              set_caption("Tukey HSD Post-hoc Comparisons") %>%
              autofit()
            
            doc <- doc %>% body_add_flextable(ft_posthoc)
          } else {
            posthoc_text <- paste("Post-hoc method:", input$anova_posthoc_method)
            doc <- doc %>% body_add_par(posthoc_text, style = "Normal")
          }
        }
        
        # Assumptions check if available
        if (!is.null(results$assumptions)) {
          doc <- doc %>%
            body_add_par("ASSUMPTIONS CHECK", style = "heading 2")
          
          assumptions_df <- data.frame(
            Test = c("Normality (Shapiro-Wilk)", "Homogeneity of Variance (Levene)"),
            Statistic = c(
              round(results$assumptions$normality$statistic, 4),
              round(results$assumptions$homogeneity$`F value`[1], 4)
            ),
            P_Value = c(
              round(results$assumptions$normality$p.value, 4),
              round(results$assumptions$homogeneity$`Pr(>F)`[1], 4)
            ),
            Interpretation = c(
              ifelse(results$assumptions$normality$p.value > 0.05, "Assumption met", "Assumption violated"),
              ifelse(results$assumptions$homogeneity$`Pr(>F)`[1] > 0.05, "Assumption met", "Assumption violated")
            )
          )
          
          ft_assumptions <- flextable(assumptions_df) %>%
            theme_zebra() %>%
            autofit()
          
          doc <- doc %>% body_add_flextable(ft_assumptions)
        }
        
        # Interpretation
        doc <- doc %>%
          body_add_par("INTERPRETATION", style = "heading 2")
        
        p_value <- results$summary[[1]]$`Pr(>F)`[1]
        f_value <- results$summary[[1]]$`F value`[1]
        
        interpretation <- if (p_value < 0.05) {
          paste("The one-way ANOVA revealed a statistically significant difference between groups ",
                "(F =", round(f_value, 3), ", p =", round(p_value, 4), "). ",
                "This suggests that there are significant differences in", input$anova_outcome, 
                "across the levels of", input$anova_group, ".")
        } else {
          paste("The one-way ANOVA did not reveal a statistically significant difference between groups ",
                "(F =", round(f_value, 3), ", p =", round(p_value, 4), "). ",
                "This suggests that there are no significant differences in", input$anova_outcome, 
                "across the levels of", input$anova_group, ".")
        }
        
        doc <- doc %>% body_add_par(interpretation, style = "Normal")
      }
      
      print(doc, target = file)
    }
  )
  
  # Two-Way ANOVA Download
  output$download_anova_2way <- downloadHandler(
    filename = function() {
      paste("twoway_anova_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(anova_2way_results())
      results <- anova_2way_results()
      
      doc <- read_docx()
      
      if (!results$success) {
        doc <- doc %>%
          body_add_par("TWO-WAY ANOVA ANALYSIS - ERROR", style = "heading 1") %>%
          body_add_par(results$error, style = "Normal")
      } else {
        # Title and overview
        doc <- doc %>%
          body_add_par("TWO-WAY ANOVA ANALYSIS REPORT", style = "heading 1") %>%
          body_add_par(paste("Outcome Variable:", input$anova_outcome_2way), style = "Normal") %>%
          body_add_par(paste("Factor 1:", input$anova_factor1), style = "Normal") %>%
          body_add_par(paste("Factor 2:", input$anova_factor2), style = "Normal") %>%
          body_add_par(paste("Interaction Term:", ifelse(input$anova_interaction, "Included", "Excluded")), style = "Normal") %>%
          body_add_par("", style = "Normal")
        
        # ANOVA results table
        doc <- doc %>%
          body_add_par("TWO-WAY ANOVA RESULTS", style = "heading 2")
        
        anova_df <- as.data.frame(results$summary[[1]])
        anova_df <- cbind(Source = rownames(anova_df), anova_df)
        rownames(anova_df) <- NULL
        
        ft_anova <- flextable(anova_df) %>%
          theme_zebra() %>%
          set_caption("Two-Way ANOVA Table") %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_anova)
        
        # Simple effects if available
        if (!is.null(results$simple_effects)) {
          doc <- doc %>%
            body_add_par("SIMPLE EFFECTS ANALYSIS", style = "heading 2")
          
          simple_effects_text <- capture.output(print(results$simple_effects))
          for (line in simple_effects_text) {
            doc <- doc %>% body_add_par(line, style = "Normal")
          }
        }
        
        # Interpretation
        doc <- doc %>%
          body_add_par("INTERPRETATION", style = "heading 2")
        
        # Extract main effects and interaction p-values
        p_values <- results$summary[[1]]$`Pr(>F)`
        sources <- rownames(results$summary[[1]])
        
        interpretation <- "Key findings:\n"
        
        for (i in 1:length(sources)) {
          if (!is.na(p_values[i])) {
            significance <- ifelse(p_values[i] < 0.05, "significant", "not significant")
            interpretation <- paste0(interpretation, 
                                     "- ", sources[i], ": ", significance, " (p = ", round(p_values[i], 4), ")\n")
          }
        }
        
        doc <- doc %>% body_add_par(interpretation, style = "Normal")
      }
      
      print(doc, target = file)
    }
  )
  
  # Welch ANOVA Download
  output$download_welch_anova <- downloadHandler(
    filename = function() {
      paste("welch_anova_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(welch_anova_results())
      results <- welch_anova_results()
      
      doc <- read_docx()
      
      if (!results$success) {
        doc <- doc %>%
          body_add_par("WELCH ANOVA ANALYSIS - ERROR", style = "heading 1") %>%
          body_add_par(results$error, style = "Normal")
      } else {
        # Title and overview
        doc <- doc %>%
          body_add_par("WELCH ANOVA ANALYSIS REPORT", style = "heading 1") %>%
          body_add_par(paste("Outcome Variable:", input$welch_outcome), style = "Normal") %>%
          body_add_par(paste("Grouping Variable:", input$welch_group), style = "Normal") %>%
          body_add_par("Method: Welch ANOVA (robust to unequal variances)", style = "Normal") %>%
          body_add_par("", style = "Normal")
        
        # Welch ANOVA results
        doc <- doc %>%
          body_add_par("WELCH ANOVA RESULTS", style = "heading 2")
        
        welch_df <- data.frame(
          Statistic = c("F value", "Numerator DF", "Denominator DF", "P-value"),
          Value = c(
            round(results$welch_test$statistic, 4),
            round(results$welch_test$parameter[1], 2),
            round(results$welch_test$parameter[2], 2),
            round(results$welch_test$p.value, 4)
          )
        )
        
        ft_welch <- flextable(welch_df) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_welch)
        
        # Games-Howell post-hoc if available
        if (!is.null(results$posthoc)) {
          doc <- doc %>%
            body_add_par("GAMES-HOWELL POST-HOC TEST", style = "heading 2")
          
          posthoc_text <- capture.output(print(results$posthoc))
          for (line in posthoc_text) {
            doc <- doc %>% body_add_par(line, style = "Normal")
          }
        }
        
        # Interpretation
        doc <- doc %>%
          body_add_par("INTERPRETATION", style = "heading 2")
        
        p_value <- results$welch_test$p.value
        f_value <- results$welch_test$statistic
        
        interpretation <- if (p_value < 0.05) {
          paste("The Welch ANOVA revealed a statistically significant difference between groups ",
                "(F =", round(f_value, 3), ", p =", round(p_value, 4), "). ",
                "This robust test confirms significant differences while accounting for potential unequal variances.")
        } else {
          paste("The Welch ANOVA did not reveal a statistically significant difference between groups ",
                "(F =", round(f_value, 3), ", p =", round(p_value, 4), ").")
        }
        
        doc <- doc %>% body_add_par(interpretation, style = "Normal")
      }
      
      print(doc, target = file)
    }
  )
  
  # One-Sample T-Test Download
  output$download_ttest_onesample <- downloadHandler(
    filename = function() {
      paste("onesample_ttest_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(ttest_onesample_results())
      results <- ttest_onesample_results()
      
      doc <- read_docx()
      
      if (!results$success) {
        doc <- doc %>%
          body_add_par("ONE-SAMPLE T-TEST - ERROR", style = "heading 1") %>%
          body_add_par(results$error, style = "Normal")
      } else {
        # Title and overview
        doc <- doc %>%
          body_add_par("ONE-SAMPLE T-TEST ANALYSIS REPORT", style = "heading 1") %>%
          body_add_par(paste("Variable:", input$ttest_onesample_var), style = "Normal") %>%
          body_add_par(paste("Test Value (μ₀):", input$ttest_onesample_mu), style = "Normal") %>%
          body_add_par(paste("Alternative Hypothesis:", input$ttest_onesample_alternative), style = "Normal") %>%
          body_add_par(paste("Confidence Level:", input$ttest_onesample_conf * 100, "%"), style = "Normal") %>%
          body_add_par("", style = "Normal")
        
        # T-test results
        doc <- doc %>%
          body_add_par("T-TEST RESULTS", style = "heading 2")
        
        ttest_df <- data.frame(
          Statistic = c("t value", "Degrees of Freedom", "P-value", 
                        "Sample Mean", "Test Value", "Mean Difference",
                        "95% CI Lower", "95% CI Upper"),
          Value = c(
            round(results$ttest$statistic, 4),
            round(results$ttest$parameter, 2),
            round(results$ttest$p.value, 4),
            round(mean(results$data), 4),
            input$ttest_onesample_mu,
            round(mean(results$data) - input$ttest_onesample_mu, 4),
            round(results$ttest$conf.int[1], 4),
            round(results$ttest$conf.int[2], 4)
          )
        )
        
        ft_ttest <- flextable(ttest_df) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_ttest)
        
        # Descriptive statistics
        doc <- doc %>%
          body_add_par("DESCRIPTIVE STATISTICS", style = "heading 2")
        
        desc_df <- data.frame(
          Statistic = c("N", "Mean", "Standard Deviation", "Standard Error", 
                        "Minimum", "Maximum", "Range"),
          Value = c(
            results$descriptives$N,
            round(results$descriptives$Mean, 4),
            round(results$descriptives$SD, 4),
            round(results$descriptives$SE, 4),
            round(min(results$data), 4),
            round(max(results$data), 4),
            round(max(results$data) - min(results$data), 4)
          )
        )
        
        ft_desc <- flextable(desc_df) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_desc)
        
        # Interpretation
        doc <- doc %>%
          body_add_par("INTERPRETATION", style = "heading 2")
        
        p_value <- results$ttest$p.value
        mean_diff <- mean(results$data) - input$ttest_onesample_mu
        
        interpretation <- if (p_value < 0.05) {
          paste("The one-sample t-test revealed a statistically significant difference ",
                "between the sample mean (", round(mean(results$data), 3), ") ",
                "and the test value (", input$ttest_onesample_mu, "), ",
                "t(", round(results$ttest$parameter, 1), ") = ", round(results$ttest$statistic, 3), 
                ", p = ", round(p_value, 4), ". ",
                "The mean difference is ", round(mean_diff, 3), ".",
                sep = "")
        } else {
          paste("The one-sample t-test did not reveal a statistically significant difference ",
                "between the sample mean (", round(mean(results$data), 3), ") ",
                "and the test value (", input$ttest_onesample_mu, "), ",
                "t(", round(results$ttest$parameter, 1), ") = ", round(results$ttest$statistic, 3), 
                ", p = ", round(p_value, 4), ". ",
                "The mean difference is ", round(mean_diff, 3), ".",
                sep = "")
        }
        
        doc <- doc %>% body_add_par(interpretation, style = "Normal")
      }
      
      print(doc, target = file)
    }
  )
  
  # Independent T-Test Download
  output$download_ttest_independent <- downloadHandler(
    filename = function() {
      paste("independent_ttest_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(ttest_independent_results())
      results <- ttest_independent_results()
      
      doc <- read_docx()
      
      if (!results$success) {
        doc <- doc %>%
          body_add_par("INDEPENDENT T-TEST - ERROR", style = "heading 1") %>%
          body_add_par(results$error, style = "Normal")
      } else {
        # Title and overview
        doc <- doc %>%
          body_add_par("INDEPENDENT T-TEST ANALYSIS REPORT", style = "heading 1") %>%
          body_add_par(paste("Outcome Variable:", input$ttest_independent_var), style = "Normal") %>%
          body_add_par(paste("Grouping Variable:", input$ttest_independent_group), style = "Normal") %>%
          body_add_par(paste("Group 1:", input$ttest_group1), style = "Normal") %>%
          body_add_par(paste("Group 2:", input$ttest_group2), style = "Normal") %>%
          body_add_par(paste("Equal Variances Assumed:", ifelse(input$ttest_var_equal, "Yes", "No")), style = "Normal") %>%
          body_add_par(paste("Alternative Hypothesis:", input$ttest_alternative), style = "Normal") %>%
          body_add_par("", style = "Normal")
        
        # T-test results
        doc <- doc %>%
          body_add_par("T-TEST RESULTS", style = "heading 2")
        
        ttest_df <- data.frame(
          Statistic = c("t value", "Degrees of Freedom", "P-value", 
                        "Mean Difference", "Standard Error", "95% CI Lower", "95% CI Upper"),
          Value = c(
            round(results$ttest$statistic, 4),
            round(results$ttest$parameter, 2),
            round(results$ttest$p.value, 4),
            round(results$ttest$estimate[1] - results$ttest$estimate[2], 4),
            round(results$ttest$stderr, 4),
            round(results$ttest$conf.int[1], 4),
            round(results$ttest$conf.int[2], 4)
          )
        )
        
        ft_ttest <- flextable(ttest_df) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_ttest)
        
        # Group descriptives
        doc <- doc %>%
          body_add_par("GROUP DESCRIPTIVES", style = "heading 2")
        
        ft_desc <- flextable(results$descriptives) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_desc)
        
        # Variance test results
        doc <- doc %>%
          body_add_par("VARIANCE EQUALITY TEST", style = "heading 2")
        
        var_df <- data.frame(
          Statistic = c("F value", "Numerator DF", "Denominator DF", "P-value", "Interpretation"),
          Value = c(
            round(results$var_test$statistic, 4),
            round(results$var_test$parameter[1], 2),
            round(results$var_test$parameter[2], 2),
            round(results$var_test$p.value, 4),
            ifelse(results$var_test$p.value < 0.05, 
                   "Variances significantly different - Welch test recommended", 
                   "Variances not significantly different - equal variances assumption reasonable")
          )
        )
        
        ft_var <- flextable(var_df) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_var)
        
        # Interpretation
        doc <- doc %>%
          body_add_par("INTERPRETATION", style = "heading 2")
        
        p_value <- results$ttest$p.value
        mean_diff <- results$ttest$estimate[1] - results$ttest$estimate[2]
        group1_mean <- results$descriptives$Mean[1]
        group2_mean <- results$descriptives$Mean[2]
        
        interpretation <- if (p_value < 0.05) {
          paste("The independent t-test revealed a statistically significant difference ",
                "between ", input$ttest_group1, " (M = ", round(group1_mean, 3), ") ",
                "and ", input$ttest_group2, " (M = ", round(group2_mean, 3), "), ",
                "t(", round(results$ttest$parameter, 1), ") = ", round(abs(results$ttest$statistic), 3), 
                ", p = ", round(p_value, 4), ". ",
                "The mean difference is ", round(mean_diff, 3), " ",
                "with a ", input$ttest_conf * 100, "% CI of [", 
                round(results$ttest$conf.int[1], 3), ", ", 
                round(results$ttest$conf.int[2], 3), "].",
                sep = "")
        } else {
          paste("The independent t-test did not reveal a statistically significant difference ",
                "between ", input$ttest_group1, " (M = ", round(group1_mean, 3), ") ",
                "and ", input$ttest_group2, " (M = ", round(group2_mean, 3), "), ",
                "t(", round(results$ttest$parameter, 1), ") = ", round(abs(results$ttest$statistic), 3), 
                ", p = ", round(p_value, 4), ". ",
                "The mean difference is ", round(mean_diff, 3), " ",
                "with a ", input$ttest_conf * 100, "% CI of [", 
                round(results$ttest$conf.int[1], 3), ", ", 
                round(results$ttest$conf.int[2], 3), "].",
                sep = "")
        }
        
        doc <- doc %>% body_add_par(interpretation, style = "Normal")
      }
      
      print(doc, target = file)
    }
  )
  
  # Paired T-Test Download
  output$download_ttest_paired <- downloadHandler(
    filename = function() {
      paste("paired_ttest_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(ttest_paired_results())
      results <- ttest_paired_results()
      
      doc <- read_docx()
      
      if (!results$success) {
        doc <- doc %>%
          body_add_par("PAIRED T-TEST - ERROR", style = "heading 1") %>%
          body_add_par(results$error, style = "Normal")
      } else {
        # Title and overview
        doc <- doc %>%
          body_add_par("PAIRED T-TEST ANALYSIS REPORT", style = "heading 1") %>%
          body_add_par(paste("Variable 1 (Before/Time 1):", input$ttest_paired_var1), style = "Normal") %>%
          body_add_par(paste("Variable 2 (After/Time 2):", input$ttest_paired_var2), style = "Normal") %>%
          body_add_par(paste("Number of Pairs:", length(results$var1_data)), style = "Normal") %>%
          body_add_par(paste("Alternative Hypothesis:", input$ttest_paired_alternative), style = "Normal") %>%
          body_add_par("", style = "Normal")
        
        # T-test results
        doc <- doc %>%
          body_add_par("PAIRED T-TEST RESULTS", style = "heading 2")
        
        ttest_df <- data.frame(
          Statistic = c("t value", "Degrees of Freedom", "P-value", 
                        "Mean Difference", "Standard Error", "95% CI Lower", "95% CI Upper"),
          Value = c(
            round(results$ttest$statistic, 4),
            round(results$ttest$parameter, 2),
            round(results$ttest$p.value, 4),
            round(results$ttest$estimate, 4),
            round(results$ttest$stderr, 4),
            round(results$ttest$conf.int[1], 4),
            round(results$ttest$conf.int[2], 4)
          )
        )
        
        ft_ttest <- flextable(ttest_df) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_ttest)
        
        # Descriptive statistics
        doc <- doc %>%
          body_add_par("DESCRIPTIVE STATISTICS", style = "heading 2")
        
        ft_desc <- flextable(results$descriptives) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_desc)
        
        # Difference analysis
        doc <- doc %>%
          body_add_par("DIFFERENCE ANALYSIS", style = "heading 2")
        
        correlation <- cor(results$var1_data, results$var2_data)
        diff_df <- data.frame(
          Statistic = c("Correlation between pairs", "Mean of differences", 
                        "SD of differences", "Minimum difference", "Maximum difference"),
          Value = c(
            round(correlation, 4),
            round(mean(results$differences), 4),
            round(sd(results$differences), 4),
            round(min(results$differences), 4),
            round(max(results$differences), 4)
          )
        )
        
        ft_diff <- flextable(diff_df) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_diff)
        
        # Interpretation
        doc <- doc %>%
          body_add_par("INTERPRETATION", style = "heading 2")
        
        p_value <- results$ttest$p.value
        mean_diff <- results$ttest$estimate
        var1_mean <- results$descriptives$Mean[1]
        var2_mean <- results$descriptives$Mean[2]
        
        interpretation <- if (p_value < 0.05) {
          paste("The paired t-test revealed a statistically significant difference ",
                "between ", input$ttest_paired_var1, " (M = ", round(var1_mean, 3), ") ",
                "and ", input$ttest_paired_var2, " (M = ", round(var2_mean, 3), "), ",
                "t(", round(results$ttest$parameter, 1), ") = ", round(abs(results$ttest$statistic), 3), 
                ", p = ", round(p_value, 4), ". ",
                "The mean difference is ", round(mean_diff, 3), " ",
                "with a correlation of ", round(correlation, 3), " between the paired measurements.",
                sep = "")
        } else {
          paste("The paired t-test did not reveal a statistically significant difference ",
                "between ", input$ttest_paired_var1, " (M = ", round(var1_mean, 3), ") ",
                "and ", input$ttest_paired_var2, " (M = ", round(var2_mean, 3), "), ",
                "t(", round(results$ttest$parameter, 1), ") = ", round(abs(results$ttest$statistic), 3), 
                ", p = ", round(p_value, 4), ". ",
                "The mean difference is ", round(mean_diff, 3), " ",
                "with a correlation of ", round(correlation, 3), " between the paired measurements.",
                sep = "")
        }
        
        doc <- doc %>% body_add_par(interpretation, style = "Normal")
      }
      
      print(doc, target = file)
    }
  )
  
  # Non-parametric Tests Download
  output$download_nonparametric <- downloadHandler(
    filename = function() {
      paste("nonparametric_tests_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(mannwhitney_results(), wilcoxon_results())
      
      doc <- read_docx() %>%
        body_add_par("NON-PARAMETRIC TESTS REPORT", style = "heading 1") %>%
        body_add_par("", style = "Normal")
      
      # Mann-Whitney U Test results
      mann_results <- mannwhitney_results()
      if (mann_results$success) {
        doc <- doc %>%
          body_add_par("MANN-WHITNEY U TEST", style = "heading 2") %>%
          body_add_par(paste("Variable:", input$mannwhitney_var), style = "Normal") %>%
          body_add_par(paste("Grouping Variable:", input$mannwhitney_group), style = "Normal") %>%
          body_add_par("", style = "Normal")
        
        mann_df <- data.frame(
          Statistic = c("W statistic", "P-value", "Alternative Hypothesis"),
          Value = c(
            round(mann_results$test$statistic, 4),
            round(mann_results$test$p.value, 4),
            mann_results$test$alternative
          )
        )
        
        ft_mann <- flextable(mann_df) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_mann)
      }
      
      # Wilcoxon Signed-Rank Test results
      wilcox_results <- wilcoxon_results()
      if (wilcox_results$success) {
        doc <- doc %>%
          body_add_par("WILCOXON SIGNED-RANK TEST", style = "heading 2") %>%
          body_add_par(paste("Variable 1:", input$wilcoxon_var1), style = "Normal") %>%
          body_add_par(paste("Variable 2:", input$wilcoxon_var2), style = "Normal") %>%
          body_add_par("", style = "Normal")
        
        wilcox_df <- data.frame(
          Statistic = c("V statistic", "P-value", "Alternative Hypothesis"),
          Value = c(
            round(wilcox_results$test$statistic, 4),
            round(wilcox_results$test$p.value, 4),
            wilcox_results$test$alternative
          )
        )
        
        ft_wilcox <- flextable(wilcox_df) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_wilcox)
      }
      
      print(doc, target = file)
    }
  )
  
  # ANOVA Assumptions Download
  output$download_assumptions <- downloadHandler(
    filename = function() {
      paste("anova_assumptions_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(assumptions_results())
      results <- assumptions_results()
      
      doc <- read_docx()
      
      if (!results$success) {
        doc <- doc %>%
          body_add_par("ANOVA ASSUMPTIONS CHECK - ERROR", style = "heading 1") %>%
          body_add_par(results$error, style = "Normal")
      } else {
        doc <- doc %>%
          body_add_par("ANOVA ASSUMPTIONS CHECK REPORT", style = "heading 1") %>%
          body_add_par(paste("Outcome Variable:", input$assumptions_outcome), style = "Normal") %>%
          body_add_par(paste("Grouping Variable:", input$assumptions_group), style = "Normal") %>%
          body_add_par("", style = "Normal")
        
        # Levene's test results
        if (!is.null(results$levene)) {
          doc <- doc %>%
            body_add_par("HOMOGENEITY OF VARIANCE", style = "heading 2")
          
          levene_df <- data.frame(
            Statistic = c("F value", "DF1", "DF2", "P-value", "Interpretation"),
            Value = c(
              round(results$levene$`F value`[1], 4),
              round(results$levene$Df[1], 0),
              round(results$levene$Df[2], 0),
              round(results$levene$`Pr(>F)`[1], 4),
              ifelse(results$levene$`Pr(>F)`[1] > 0.05, 
                     "Assumption met - variances are homogeneous", 
                     "Assumption violated - variances are not homogeneous")
            )
          )
          
          ft_levene <- flextable(levene_df) %>%
            theme_zebra() %>%
            autofit()
          
          doc <- doc %>% body_add_flextable(ft_levene)
        }
        
        # Normality tests
        if (!is.null(results$shapiro)) {
          doc <- doc %>%
            body_add_par("NORMALITY TESTS", style = "heading 2")
          
          normality_df <- data.frame(
            Test = c("Shapiro-Wilk", "Kolmogorov-Smirnov"),
            Statistic = c(
              round(results$shapiro$statistic, 4),
              ifelse(!is.null(results$ks), round(results$ks$statistic, 4), "N/A")
            ),
            P_Value = c(
              round(results$shapiro$p.value, 4),
              ifelse(!is.null(results$ks), round(results$ks$p.value, 4), "N/A")
            ),
            Interpretation = c(
              ifelse(results$shapiro$p.value > 0.05, "Normal", "Non-normal"),
              ifelse(!is.null(results$ks) && results$ks$p.value > 0.05, "Normal", "Non-normal")
            )
          )
          
          ft_normality <- flextable(normality_df) %>%
            theme_zebra() %>%
            autofit()
          
          doc <- doc %>% body_add_flextable(ft_normality)
        }
        
        # Outlier detection
        if (!is.null(results$outliers)) {
          doc <- doc %>%
            body_add_par("OUTLIER DETECTION", style = "heading 2")
          
          outlier_text <- if (length(results$outliers) > 0) {
            paste("Outliers detected at observations:", paste(results$outliers, collapse = ", "))
          } else {
            "No significant outliers detected (standardized residuals < |2.5|)"
          }
          
          doc <- doc %>% body_add_par(outlier_text, style = "Normal")
        }
        
        # Overall assessment
        doc <- doc %>%
          body_add_par("OVERALL ASSESSMENT", style = "heading 2")
        
        assumptions_met <- TRUE
        issues <- c()
        
        if (!is.null(results$levene) && results$levene$`Pr(>F)`[1] < 0.05) {
          assumptions_met <- FALSE
          issues <- c(issues, "Heterogeneous variances detected")
        }
        
        if (!is.null(results$shapiro) && results$shapiro$p.value < 0.05) {
          assumptions_met <- FALSE
          issues <- c(issues, "Non-normal residuals detected")
        }
        
        if (!is.null(results$outliers) && length(results$outliers) > 0) {
          issues <- c(issues, paste(length(results$outliers), "outliers detected"))
        }
        
        assessment <- if (assumptions_met && length(issues) == 0) {
          "All key ANOVA assumptions appear to be reasonably met. Results should be valid."
        } else {
          paste("Potential issues detected:", paste(issues, collapse = "; "), 
                ". Consider using robust alternatives or data transformations.")
        }
        
        doc <- doc %>% body_add_par(assessment, style = "Normal")
      }
      
      print(doc, target = file)
    }
  )
  # Server function - CORRECTED RISK RATIO SECTION
  
  # UI for outcome and exposure level selection
  output$rr_outcome_levels_ui <- renderUI({
    req(input$rr_outcome, processed_data())
    df <- processed_data()
    
    if (!input$rr_outcome %in% names(df)) return(NULL)
    
    var_data <- df[[input$rr_outcome]]
    levels <- if (is.factor(var_data)) levels(var_data) else unique(na.omit(var_data))
    
    if (length(levels) > 10) {
      return(helpText("Outcome variable has too many categories for risk ratio calculation."))
    }
    
    tagList(
      selectInput("rr_outcome_ref", "Outcome Reference Level (Baseline):", choices = levels),
      selectInput("rr_outcome_comp", "Outcome Comparison Level:", choices = levels, selected = levels[2])
    )
  })
  
  output$rr_exposure_levels_ui <- renderUI({
    req(input$rr_exposure, processed_data())
    df <- processed_data()
    
    if (!input$rr_exposure %in% names(df)) return(NULL)
    
    var_data <- df[[input$rr_exposure]]
    levels <- if (is.factor(var_data)) levels(var_data) else unique(na.omit(var_data))
    
    if (length(levels) > 10) {
      return(helpText("Exposure variable has too many categories for risk ratio calculation."))
    }
    
    tagList(
      selectInput("rr_exposure_ref", "Exposure Reference Level (Unexposed):", choices = levels),
      selectInput("rr_exposure_comp", "Exposure Comparison Level (Exposed):", choices = levels, selected = levels[2])
    )
  })
  
  # ULTRA-ROBUST RISK RATIO ANALYSIS - COMPLETELY REWRITTEN
  rr_results <- eventReactive(input$run_rr, {
    req(input$rr_outcome, input$rr_exposure, processed_data())
    req(input$rr_outcome_ref, input$rr_outcome_comp, input$rr_exposure_ref, input$rr_exposure_comp)
    
    tryCatch({
      df <- processed_data()
      
      # Validate that variables exist
      if (!input$rr_outcome %in% names(df) || !input$rr_exposure %in% names(df)) {
        return(list(
          error = "Selected variables not found in dataset",
          success = FALSE
        ))
      }
      
      # Get the variables
      outcome_var <- df[[input$rr_outcome]]
      exposure_var <- df[[input$rr_exposure]]
      
      # Remove missing values
      complete_cases <- !is.na(outcome_var) & !is.na(exposure_var)
      outcome_clean <- outcome_var[complete_cases]
      exposure_clean <- exposure_var[complete_cases]
      
      if (length(outcome_clean) == 0 || length(exposure_clean) == 0) {
        return(list(
          error = "No complete cases after removing missing values",
          success = FALSE
        ))
      }
      
      # Convert to factors with user-specified levels
      outcome_factor <- factor(outcome_clean)
      exposure_factor <- factor(exposure_clean)
      
      # Ensure user-selected levels exist in the data
      if (!input$rr_outcome_ref %in% levels(outcome_factor) || 
          !input$rr_outcome_comp %in% levels(outcome_factor)) {
        return(list(
          error = "Selected outcome levels not found in data",
          success = FALSE
        ))
      }
      
      if (!input$rr_exposure_ref %in% levels(exposure_factor) || 
          !input$rr_exposure_comp %in% levels(exposure_factor)) {
        return(list(
          error = "Selected exposure levels not found in data",
          success = FALSE
        ))
      }
      
      # Relevel factors based on user selection
      outcome_factor <- relevel(outcome_factor, ref = input$rr_outcome_ref)
      exposure_factor <- relevel(exposure_factor, ref = input$rr_exposure_ref)
      
      # Create 2x2 contingency table
      tab <- table(
        Exposure = exposure_factor,
        Outcome = outcome_factor
      )
      
      # Keep only the two selected levels for each variable
      tab_2x2 <- tab[
        c(input$rr_exposure_ref, input$rr_exposure_comp),
        c(input$rr_outcome_ref, input$rr_outcome_comp)
      ]
      
      # Check for zero cells
      if (any(tab_2x2 == 0)) {
        # Apply Haldane-Anscombe correction for zero cells
        tab_2x2_corrected <- tab_2x2 + 0.5
        correction_applied <- TRUE
      } else {
        tab_2x2_corrected <- tab_2x2
        correction_applied <- FALSE
      }
      
      # Extract values from the 2x2 table
      a <- tab_2x2_corrected[2, 2]  # Exposed cases
      b <- tab_2x2_corrected[2, 1]  # Exposed non-cases
      c <- tab_2x2_corrected[1, 2]  # Unexposed cases
      d <- tab_2x2_corrected[1, 1]  # Unexposed non-cases
      
      # Calculate risks
      risk_exposed <- a / (a + b)
      risk_unexposed <- c / (c + d)
      
      # Calculate risk ratio
      risk_ratio <- risk_exposed / risk_unexposed
      
      # Calculate confidence intervals using log method
      log_rr <- log(risk_ratio)
      se_log_rr <- sqrt((1/a) - (1/(a+b)) + (1/c) - (1/(c+d)))
      
      z <- qnorm(1 - (1 - input$rr_conf) / 2)
      ci_lower <- exp(log_rr - z * se_log_rr)
      ci_upper <- exp(log_rr + z * se_log_rr)
      
      # Calculate risk difference
      risk_diff <- risk_exposed - risk_unexposed
      se_rd <- sqrt((risk_exposed * (1 - risk_exposed) / (a + b)) + 
                      (risk_unexposed * (1 - risk_unexposed) / (c + d)))
      rd_ci_lower <- risk_diff - z * se_rd
      rd_ci_upper <- risk_diff + z * se_rd
      
      # Create comprehensive results table
      results_table <- data.frame(
        Statistic = c(
          "Exposed Group",
          "Unexposed Group", 
          "Risk in Exposed",
          "Risk in Unexposed",
          "Risk Ratio (RR)",
          paste("RR Lower CI (", input$rr_conf * 100, "%)", sep = ""),
          paste("RR Upper CI (", input$rr_conf * 100, "%)", sep = ""),
          "Risk Difference",
          paste("RD Lower CI (", input$rr_conf * 100, "%)", sep = ""),
          paste("RD Upper CI (", input$rr_conf * 100, "%)", sep = ""),
          "Exposed Cases",
          "Exposed Non-Cases",
          "Unexposed Cases", 
          "Unexposed Non-Cases",
          "Total Exposed",
          "Total Unexposed",
          "Total Cases",
          "Total Observations",
          "Zero Cell Correction"
        ),
        Value = c(
          input$rr_exposure_comp,
          input$rr_exposure_ref,
          round(risk_exposed, 4),
          round(risk_unexposed, 4),
          round(risk_ratio, 4),
          round(ci_lower, 4),
          round(ci_upper, 4),
          round(risk_diff, 4),
          round(rd_ci_lower, 4),
          round(rd_ci_upper, 4),
          a, b, c, d,
          a + b,
          c + d,
          a + c,
          a + b + c + d,
          ifelse(correction_applied, "Yes (Haldane-Anscombe)", "No")
        ),
        stringsAsFactors = FALSE
      )
      
      # Create contingency table for display
      contingency_df <- as.data.frame.matrix(tab_2x2)
      colnames(contingency_df) <- paste("Outcome:", colnames(contingency_df))
      rownames(contingency_df) <- paste("Exposure:", rownames(contingency_df))
      
      # Add totals
      contingency_df$Total <- rowSums(contingency_df)
      total_row <- c(colSums(contingency_df[, -ncol(contingency_df), drop = FALSE]), 
                     sum(contingency_df$Total))
      contingency_df <- rbind(contingency_df, Total = total_row)
      
      list(
        table = results_table,
        contingency = contingency_df,
        raw_table = tab_2x2,
        risk_ratio = risk_ratio,
        ci_lower = ci_lower,
        ci_upper = ci_upper,
        risk_exposed = risk_exposed,
        risk_unexposed = risk_unexposed,
        correction_applied = correction_applied,
        success = TRUE
      )
      
    }, error = function(e) {
      return(list(
        error = paste("Risk ratio calculation failed:", e$message),
        success = FALSE
      ))
    })
  })
  
  # Output for risk ratio results
  output$rr_results <- renderPrint({
    req(rr_results())
    
    results <- rr_results()
    
    if (!results$success) {
      cat("RISK RATIO ANALYSIS - ERROR\n")
      cat("===========================\n")
      cat(results$error, "\n")
      return()
    }
    
    cat("RISK RATIO ANALYSIS\n")
    cat("===================\n\n")
    
    cat("CONTINGENCY TABLE:\n")
    print(results$raw_table)
    cat("\n")
    
    cat("RISK RATIO RESULTS:\n")
    print(results$table)
  })
  
  # Additional interpretation output
  output$rr_interpretation <- renderPrint({
    req(rr_results())
    
    results <- rr_results()
    
    if (!results$success) return()
    
    cat("\nINTERPRETATION:\n")
    cat("===============\n\n")
    
    rr <- results$risk_ratio
    ci_low <- results$ci_lower
    ci_high <- results$ci_upper
    
    if (rr > 1) {
      cat(sprintf("• The %s group has %.2f times the risk of %s compared to the %s group.\n", 
                  input$rr_exposure_comp, rr, input$rr_outcome_comp, input$rr_exposure_ref))
    } else if (rr < 1) {
      cat(sprintf("• The %s group has %.2f times the risk of %s compared to the %s group.\n", 
                  input$rr_exposure_comp, rr, input$rr_outcome_comp, input$rr_exposure_ref))
      cat(sprintf("  (This represents a %.1f%% reduction in risk)\n", (1 - rr) * 100))
    } else {
      cat(sprintf("• There is no difference in risk between the %s and %s groups.\n", 
                  input$rr_exposure_comp, input$rr_exposure_ref))
    }
    
    cat(sprintf("• We are %d%% confident that the true risk ratio lies between %.2f and %.2f.\n", 
                input$rr_conf * 100, ci_low, ci_high))
    
    if (ci_low > 1 || ci_high < 1) {
      cat("• This result is statistically significant (confidence interval does not include 1).\n")
    } else {
      cat("• This result is not statistically significant (confidence interval includes 1).\n")
    }
    
    if (results$correction_applied) {
      cat("• Note: Haldane-Anscombe correction was applied due to zero cells in the contingency table.\n")
    }
    
    # Public health significance
    risk_diff <- results$risk_exposed - results$risk_unexposed
    if (abs(risk_diff) > 0.1) {
      cat(sprintf("• This represents an absolute risk difference of %.1f%%.\n", risk_diff * 100))
    }
  })
  
  output$rr_table <- renderDT({
    req(rr_results())
    
    results <- rr_results()
    
    if (!results$success) {
      datatable(
        data.frame(Error = results$error),
        options = list(dom = 't'),
        rownames = FALSE,
        caption = "Risk Ratio Analysis - Error"
      )
    } else {
      datatable(
        results$table,
        options = list(
          pageLength = 20,
          dom = 'Blfrtip',
          scrollX = TRUE
        ),
        rownames = FALSE,
        caption = "Risk Ratio Analysis Results"
      )
    }
  })
  
  # Download handler for risk ratios
  output$download_rr <- downloadHandler(
    filename = function() {
      paste("risk_ratio_analysis_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(rr_results())
      
      results <- rr_results()
      
      if (!results$success) {
        # Create error document
        ft <- flextable(data.frame(Error = results$error)) %>%
          set_caption("Risk Ratio Analysis - Error") %>%
          theme_zebra()
        
        doc <- read_docx() %>% 
          body_add_flextable(ft)
      } else {
        # Create comprehensive report
        doc <- read_docx()
        
        # Title and overview
        doc <- doc %>%
          body_add_par("RISK RATIO ANALYSIS REPORT", style = "heading 1") %>%
          body_add_par(paste("Outcome Variable:", input$rr_outcome), style = "Normal") %>%
          body_add_par(paste("Exposure Variable:", input$rr_exposure), style = "Normal") %>%
          body_add_par(paste("Confidence Level:", input$rr_conf * 100, "%"), style = "Normal") %>%
          body_add_par("", style = "Normal")
        
        # Contingency table
        doc <- doc %>%
          body_add_par("CONTINGENCY TABLE", style = "heading 2")
        
        contingency_with_rownames <- cbind(
          Category = rownames(results$contingency), 
          results$contingency
        )
        ft_contingency <- flextable(contingency_with_rownames) %>%
          theme_zebra() %>%
          autofit()
        doc <- doc %>% body_add_flextable(ft_contingency)
        
        # Risk ratio results
        doc <- doc %>%
          body_add_par("RISK RATIO RESULTS", style = "heading 2")
        
        ft_results <- flextable(results$table) %>%
          theme_zebra() %>%
          autofit()
        doc <- doc %>% body_add_flextable(ft_results)
        
        # Interpretation
        doc <- doc %>%
          body_add_par("INTERPRETATION", style = "heading 2")
        
        rr <- results$risk_ratio
        interpretation_text <- if (rr > 1) {
          paste("The", input$rr_exposure_comp, "group has", round(rr, 2), 
                "times the risk of", input$rr_outcome_comp, "compared to the", 
                input$rr_exposure_ref, "group.")
        } else if (rr < 1) {
          paste("The", input$rr_exposure_comp, "group has", round(rr, 2), 
                "times the risk of", input$rr_outcome_comp, "compared to the", 
                input$rr_exposure_ref, "group, representing a", 
                round((1 - rr) * 100, 1), "% reduction in risk.")
        } else {
          paste("There is no difference in risk between the", input$rr_exposure_comp, 
                "and", input$rr_exposure_ref, "groups.")
        }
        
        doc <- doc %>% 
          body_add_par(interpretation_text, style = "Normal") %>%
          body_add_par(paste("Confidence interval:", round(results$ci_lower, 2), "to", 
                             round(results$ci_upper, 2)), style = "Normal")
      }
      
      print(doc, target = file)
    }
  )
  
  # Ultra-robust Odds Ratio Analysis - FIXED for Inf and calculation issues
  or_results <- eventReactive(input$run_or, {
    req(input$or_outcome, input$or_exposure, processed_data())
    
    tryCatch({
      df <- processed_data()
      
      # Double-check variables exist with proper NA handling
      if (is.null(input$or_outcome) || !input$or_outcome %in% names(df)) {
        return(list(
          table = data.frame(),
          contingency = data.frame(),
          error = "Outcome variable not found in dataset",
          success = FALSE
        ))
      }
      
      if (is.null(input$or_exposure) || !input$or_exposure %in% names(df)) {
        return(list(
          table = data.frame(),
          contingency = data.frame(),
          error = "Exposure variable not found in dataset",
          success = FALSE
        ))
      }
      
      # Safely get the variables
      outcome_data <- df[[input$or_outcome]]
      exposure_data <- df[[input$or_exposure]]
      
      # Handle empty or all-NA data
      if (length(outcome_data) == 0 || (all(is.na(outcome_data)) && length(outcome_data) > 0)) {
        return(list(
          table = data.frame(),
          contingency = data.frame(),
          error = "Outcome variable contains no data",
          success = FALSE
        ))
      }
      
      if (length(exposure_data) == 0 || (all(is.na(exposure_data)) && length(exposure_data) > 0)) {
        return(list(
          table = data.frame(),
          contingency = data.frame(),
          error = "Exposure variable contains no data",
          success = FALSE
        ))
      }
      
      # Remove rows with missing values in either variable
      complete_cases <- !is.na(outcome_data) & !is.na(exposure_data)
      
      if (sum(complete_cases, na.rm = TRUE) == 0) {
        return(list(
          table = data.frame(),
          contingency = data.frame(),
          error = "No complete cases after removing missing values",
          success = FALSE
        ))
      }
      
      df_clean <- df[complete_cases, ]
      
      # Convert to factors and ensure proper reference levels
      outcome_factor <- as.factor(df_clean[[input$or_outcome]])
      exposure_factor <- as.factor(df_clean[[input$or_exposure]])
      
      # Get levels for user information
      outcome_levels <- levels(outcome_factor)
      exposure_levels <- levels(exposure_factor)
      
      # Ensure we have at least 2 levels
      if (length(outcome_levels) < 2) {
        return(list(
          table = data.frame(),
          contingency = data.frame(),
          error = "Outcome variable must have at least 2 categories",
          success = FALSE
        ))
      }
      
      if (length(exposure_levels) < 2) {
        return(list(
          table = data.frame(),
          contingency = data.frame(),
          error = "Exposure variable must have at least 2 categories",
          success = FALSE
        ))
      }
      
      # Create contingency table
      contingency_table <- table(exposure_factor, outcome_factor)
      
      # Ensure we're working with a 2x2 table for odds ratio
      # If more than 2 levels, use the first two levels
      if (nrow(contingency_table) > 2) {
        # Use the two most common exposure categories
        row_sums <- rowSums(contingency_table)
        top_two <- names(sort(row_sums, decreasing = TRUE))[1:2]
        contingency_table <- contingency_table[top_two, ]
        exposure_levels <- top_two
      }
      
      if (ncol(contingency_table) > 2) {
        # Use the two most common outcome categories
        col_sums <- colSums(contingency_table)
        top_two <- names(sort(col_sums, decreasing = TRUE))[1:2]
        contingency_table <- contingency_table[, top_two]
        outcome_levels <- top_two
      }
      
      # Extract the 2x2 table values
      a <- as.numeric(contingency_table[1, 1])  # Exposed cases
      b <- as.numeric(contingency_table[1, 2])  # Exposed non-cases
      c <- as.numeric(contingency_table[2, 1])  # Unexposed cases
      d <- as.numeric(contingency_table[2, 2])  # Unexposed non-cases
      
      # Check for problematic tables that cause Inf OR
      has_zero_cells <- any(c(a, b, c, d) == 0)
      has_undefined_or <- (b == 0 && c == 0) || (a == 0 && d == 0)
      
      # Calculate odds ratio with robust zero-cell handling
      calculate_or_safe <- function(a, b, c, d, conf_level) {
        
        # Apply Haldane-Anscombe correction for zero cells
        if (has_zero_cells) {
          a_corr <- a + 0.5
          b_corr <- b + 0.5
          c_corr <- c + 0.5
          d_corr <- d + 0.5
          correction_applied <- TRUE
        } else {
          a_corr <- a
          b_corr <- b
          c_corr <- c
          d_corr <- d
          correction_applied = FALSE
        }
        
        # Calculate odds ratio
        if (b_corr == 0 && c_corr == 0) {
          # Both denominators are zero - OR is undefined
          return(list(or = Inf, ci_lower = Inf, ci_upper = Inf, method = "undefined"))
        } else if (b_corr == 0) {
          # Only b is zero - OR approaches infinity
          or_value <- Inf
        } else if (c_corr == 0) {
          # Only c is zero - OR approaches infinity
          or_value <- Inf
        } else {
          # Normal calculation
          or_value <- (a_corr * d_corr) / (b_corr * c_corr)
        }
        
        # Calculate confidence intervals if possible
        if (is.finite(or_value) && or_value > 0 && !is.infinite(or_value)) {
          log_or <- log(or_value)
          se_log_or <- sqrt(1/a_corr + 1/b_corr + 1/c_corr + 1/d_corr)
          z <- qnorm(1 - (1 - conf_level) / 2)
          
          ci_lower <- exp(log_or - z * se_log_or)
          ci_upper <- exp(log_or + z * se_log_or)
        } else {
          ci_lower <- NA
          ci_upper <- NA
        }
        
        return(list(
          or = or_value,
          ci_lower = ci_lower,
          ci_upper = ci_upper,
          method = ifelse(correction_applied, "corrected for zero cells", "standard"),
          correction_applied = correction_applied
        ))
      }
      
      # Calculate odds ratio
      or_calc <- calculate_or_safe(a, b, c, d, input$or_conf)
      
      # Create comprehensive results table
      results_table <- data.frame(
        Statistic = c(
          "Outcome Variable",
          "Outcome Levels",
          "Exposure Variable", 
          "Exposure Levels",
          "Odds Ratio",
          paste("Lower CI (", input$or_conf * 100, "%)", sep = ""),
          paste("Upper CI (", input$or_conf * 100, "%)", sep = ""),
          "Calculation Method",
          "Zero Cells Detected",
          "Correction Applied",
          "Exposed Cases",
          "Exposed Non-Cases", 
          "Unexposed Cases",
          "Unexposed Non-Cases",
          "Total Cases",
          "Total Non-Cases",
          "Total Exposed",
          "Total Unexposed",
          "Total Observations"
        ),
        Value = c(
          input$or_outcome,
          paste(outcome_levels, collapse = ", "),
          input$or_exposure,
          paste(exposure_levels, collapse = ", "),
          ifelse(is.infinite(or_calc$or), "Infinity", 
                 ifelse(is.na(or_calc$or), "Cannot calculate", round(or_calc$or, 4))),
          ifelse(is.na(or_calc$ci_lower), "Cannot calculate", round(or_calc$ci_lower, 4)),
          ifelse(is.na(or_calc$ci_upper), "Cannot calculate", round(or_calc$ci_upper, 4)),
          or_calc$method,
          ifelse(has_zero_cells, "Yes", "No"),
          ifelse(or_calc$correction_applied, "Yes", "No"),
          a, b, c, d,
          a + c,  # Total cases
          b + d,  # Total non-cases
          a + b,  # Total exposed
          c + d,  # Total unexposed
          a + b + c + d  # Total observations
        ),
        stringsAsFactors = FALSE
      )
      
      # Create formatted contingency table for display
      contingency_df <- as.data.frame.matrix(contingency_table)
      colnames(contingency_df) <- paste("Outcome:", colnames(contingency_df))
      rownames(contingency_df) <- paste("Exposure:", rownames(contingency_df))
      
      # Add row and column totals
      contingency_df$Total <- rowSums(contingency_df)
      total_row <- c(colSums(contingency_df[, -ncol(contingency_df), drop = FALSE]), 
                     sum(contingency_df$Total))
      contingency_df <- rbind(contingency_df, Total = total_row)
      
      # Calculate additional statistics
      additional_stats <- data.frame(
        Statistic = character(),
        Value = character(),
        stringsAsFactors = FALSE
      )
      
      # Only calculate chi-square if we have reasonable data
      if (sum(contingency_table) >= 20 && all(contingency_table >= 5)) {
        chi_test <- tryCatch({
          result <- chisq.test(contingency_table)
          data.frame(
            Statistic = c("Chi-square Statistic", "Chi-square P-value"),
            Value = c(round(result$statistic, 4), round(result$p.value, 4))
          )
        }, error = function(e) {
          data.frame(
            Statistic = c("Chi-square Statistic", "Chi-square P-value"),
            Value = c("Cannot calculate", "Cannot calculate")
          )
        })
        additional_stats <- rbind(additional_stats, chi_test)
      }
      
      # Fisher's exact test for small samples or when chi-square assumptions violated
      fisher_test <- tryCatch({
        result <- fisher.test(contingency_table, conf.level = input$or_conf)
        data.frame(
          Statistic = c("Fisher's Exact OR", "Fisher's Exact P-value"),
          Value = c(round(result$estimate, 4), round(result$p.value, 4))
        )
      }, error = function(e) {
        data.frame(
          Statistic = c("Fisher's Exact OR", "Fisher's Exact P-value"),
          Value = c("Cannot calculate", "Cannot calculate")
        )
      })
      additional_stats <- rbind(additional_stats, fisher_test)
      
      # Risk ratio and risk difference
      if (a + b > 0 && c + d > 0) {
        risk_exposed <- a / (a + b)
        risk_unexposed <- c / (c + d)
        risk_ratio <- risk_exposed / risk_unexposed
        risk_diff <- risk_exposed - risk_unexposed
        
        risk_stats <- data.frame(
          Statistic = c("Risk in Exposed", "Risk in Unexposed", "Risk Ratio", "Risk Difference"),
          Value = c(
            round(risk_exposed, 4),
            round(risk_unexposed, 4),
            round(risk_ratio, 4),
            round(risk_diff, 4)
          )
        )
        additional_stats <- rbind(additional_stats, risk_stats)
      }
      
      list(
        table = results_table,
        contingency = contingency_df,
        additional_stats = additional_stats,
        raw_table = contingency_table,
        or_calc = or_calc,
        success = TRUE
      )
      
    }, error = function(e) {
      return(list(
        table = data.frame(),
        contingency = data.frame(),
        error = paste("Error:", e$message),
        success = FALSE
      ))
    })
  })
  
  output$or_results <- renderPrint({
    req(or_results())
    
    results <- or_results()
    
    if (!results$success) {
      cat("ODDS RATIO ANALYSIS - ERROR\n")
      cat("===========================\n")
      cat(results$error, "\n")
      return()
    }
    
    cat("ODDS RATIO ANALYSIS\n")
    cat("===================\n\n")
    
    cat("CONTINGENCY TABLE:\n")
    print(results$raw_table)
    cat("\n")
    
    cat("ODDS RATIO RESULTS:\n")
    print(results$table)
    cat("\n")
    
    if (nrow(results$additional_stats) > 0) {
      cat("ADDITIONAL STATISTICS:\n")
      print(results$additional_stats)
      cat("\n")
    }
    
    # Enhanced interpretation
    or_value <- results$or_calc$or
    ci_lower <- results$or_calc$ci_lower
    ci_upper <- results$or_calc$ci_upper
    
    cat("INTERPRETATION:\n")
    
    if (is.infinite(or_value)) {
      cat("• The odds ratio is infinite due to zero cells in the contingency table.\n")
      cat("• This occurs when there are no cases in one of the exposure groups.\n")
      cat("• Consider using Fisher's exact test or reviewing your data categories.\n")
    } else if (is.na(or_value)) {
      cat("• Cannot calculate odds ratio due to data issues.\n")
    } else {
      # Normal interpretation
      if (or_value > 1) {
        cat(sprintf("• The odds of the outcome in the exposed group are %.2f times higher than in the unexposed group.\n", or_value))
      } else if (or_value < 1) {
        cat(sprintf("• The odds of the outcome in the exposed group are %.2f times lower than in the unexposed group.\n", 1/or_value))
      } else {
        cat("• There is no association between exposure and outcome (OR = 1).\n")
      }
      
      if (!is.na(ci_lower) && !is.na(ci_upper)) {
        if (ci_lower > 1 || ci_upper < 1) {
          cat("• The association is statistically significant (confidence interval does not include 1).\n")
        } else {
          cat("• The association is not statistically significant (confidence interval includes 1).\n")
        }
        cat(sprintf("• We are %d%% confident that the true odds ratio lies between %.2f and %.2f.\n", 
                    input$or_conf * 100, ci_lower, ci_upper))
      }
    }
    
    # Data quality notes
    if (results$or_calc$correction_applied) {
      cat("• Note: Haldane-Anscombe correction (adding 0.5 to all cells) was applied due to zero cells.\n")
    }
    
    if (any(results$raw_table == 0)) {
      cat("• Warning: Zero cells detected in the contingency table. Results should be interpreted with caution.\n")
    }
  })
  
  output$or_table <- renderDT({
    req(or_results())
    
    results <- or_results()
    
    if (!results$success) {
      # Show error in table format
      datatable(data.frame(Error = results$error), 
                options = list(dom = 't'),
                rownames = FALSE,
                caption = "Odds Ratio Analysis - Error")
    } else {
      # Show main results table
      datatable(results$table,
                options = list(
                  pageLength = 15,
                  dom = 'Blfrtip',
                  scrollX = TRUE
                ),
                rownames = FALSE,
                caption = "Odds Ratio Analysis Results")
    }
  })
  
  output$download_or <- downloadHandler(
    filename = function() {
      paste("odds_ratio_analysis_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(or_results())
      
      results <- or_results()
      
      if (!results$success) {
        # Create error document
        ft <- flextable(data.frame(Error = results$error)) %>%
          set_caption("Odds Ratio Analysis - Error") %>%
          theme_zebra()
        
        doc <- read_docx() %>% 
          body_add_flextable(ft)
      } else {
        # Create comprehensive report
        doc <- read_docx()
        
        # Add main title
        doc <- doc %>%
          body_add_par("ODDS RATIO ANALYSIS REPORT", style = "heading 1") %>%
          body_add_par(paste("Outcome:", input$or_outcome), style = "Normal") %>%
          body_add_par(paste("Exposure:", input$or_exposure), style = "Normal") %>%
          body_add_par(paste("Confidence Level:", input$or_conf * 100, "%"), style = "Normal") %>%
          body_add_par("", style = "Normal")
        
        # Add contingency table
        doc <- doc %>%
          body_add_par("CONTINGENCY TABLE", style = "heading 2")
        
        contingency_with_rownames <- cbind(Exposure = rownames(results$contingency), results$contingency)
        ft_contingency <- flextable(contingency_with_rownames) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_contingency)
        
        # Add odds ratio results
        doc <- doc %>%
          body_add_par("ODDS RATIO RESULTS", style = "heading 2")
        
        ft_main <- flextable(results$table) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_main)
        
        # Add additional statistics
        doc <- doc %>%
          body_add_par("ADDITIONAL STATISTICS", style = "heading 2")
        
        ft_additional <- flextable(results$additional_stats) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_additional)
      }
      
      print(doc, target = file)
    }
  )
  
  # Ultra-robust Kaplan-Meier Analysis - FIXED VERSION
  surv_results <- eventReactive(input$run_surv, {
    req(input$surv_time, input$surv_event, processed_data())
    
    tryCatch({
      df <- processed_data()
      
      # Double-check variables exist
      if (!input$surv_time %in% names(df)) {
        return(list(
          fit = NULL,
          error = "Time variable not found in dataset",
          success = FALSE
        ))
      }
      
      if (!input$surv_event %in% names(df)) {
        return(list(
          fit = NULL,
          error = "Event variable not found in dataset",
          success = FALSE
        ))
      }
      
      if (input$surv_group != "none" && !input$surv_group %in% names(df)) {
        return(list(
          fit = NULL,
          error = "Grouping variable not found in dataset",
          success = FALSE
        ))
      }
      
      # Safely get the variables
      time_data <- df[[input$surv_time]]
      event_data <- df[[input$surv_event]]
      
      # Handle empty or all-NA data
      if (length(time_data) == 0 || all(is.na(time_data))) {
        return(list(
          fit = NULL,
          error = "Time variable contains no data",
          success = FALSE
        ))
      }
      
      if (length(event_data) == 0 || all(is.na(event_data))) {
        return(list(
          fit = NULL,
          error = "Event variable contains no data",
          success = FALSE
        ))
      }
      
      # Convert to numeric - ROBUST CONVERSION
      time_numeric <- tryCatch({
        as.numeric(time_data)
      }, warning = function(w) {
        # If warning, try converting factor to numeric
        if (is.factor(time_data)) {
          as.numeric(as.character(time_data))
        } else {
          as.numeric(time_data)
        }
      }, error = function(e) {
        rep(NA, length(time_data))
      })
      
      event_numeric <- tryCatch({
        # First try direct conversion
        event_num <- as.numeric(event_data)
        if (all(is.na(event_num))) {
          # If that fails, convert factor/character to numeric
          if (is.factor(event_data) || is.character(event_data)) {
            event_num <- as.numeric(as.factor(event_data))
          } else {
            event_num <- as.numeric(event_data)
          }
        }
        event_num
      }, error = function(e) {
        rep(NA, length(event_data))
      })
      
      # Check for successful conversion
      if (all(is.na(time_numeric))) {
        return(list(
          fit = NULL,
          error = "Time variable could not be converted to numeric",
          success = FALSE
        ))
      }
      
      if (all(is.na(event_numeric))) {
        return(list(
          fit = NULL,
          error = "Event variable could not be converted to numeric",
          success = FALSE
        ))
      }
      
      # Remove rows with missing time or event data
      valid_cases <- !is.na(time_numeric) & !is.na(event_numeric) & time_numeric >= 0
      df_clean <- df[valid_cases, ]
      time_clean <- time_numeric[valid_cases]
      event_clean <- event_numeric[valid_cases]
      
      if (nrow(df_clean) == 0) {
        return(list(
          fit = NULL,
          error = "No valid data after removing missing/invalid time and event values",
          success = FALSE
        ))
      }
      
      # Ensure event is binary (0/1) - ROBUST APPROACH
      unique_events <- unique(event_clean)
      if (length(unique_events) == 1) {
        return(list(
          fit = NULL,
          error = "Event variable has only one unique value. Need both events and censoring.",
          success = FALSE
        ))
      }
      
      # Standardize event to 0/1 if needed
      if (!all(unique_events %in% c(0, 1))) {
        # Map to 0/1: smallest value becomes 0 (censored), others become 1 (event)
        event_clean <- ifelse(event_clean == min(unique_events), 0, 1)
      }
      
      # Check if we have both events and censoring
      if (sum(event_clean == 1) == 0) {
        return(list(
          fit = NULL,
          error = "No events found in the data (all observations are censored)",
          success = FALSE
        ))
      }
      
      # Create survival object
      surv_obj <- survival::Surv(time = time_clean, event = event_clean)
      
      # Fit Kaplan-Meier model
      if (input$surv_group != "none") {
        group_data <- df_clean[[input$surv_group]]
        if (all(is.na(group_data))) {
          return(list(
            fit = NULL,
            error = "Grouping variable contains no valid data",
            success = FALSE
          ))
        }
        
        # Convert grouping variable to factor and handle NAs
        group_var <- as.factor(group_data)
        group_var[is.na(group_var)] <- "Missing"
        
        # Check if we have enough groups
        if (length(levels(group_var)) < 2) {
          return(list(
            fit = NULL,
            error = "Grouping variable must have at least 2 categories",
            success = FALSE
          ))
        }
        
        # Create data frame for survfit
        surv_data <- data.frame(
          time = time_clean,
          event = event_clean,
          group = group_var
        )
        
        surv_fit <- survfit(Surv(time, event) ~ group, data = surv_data)
        
        # Store group information
        group_levels <- levels(group_var)
        n_groups <- length(group_levels)
      } else {
        surv_data <- data.frame(
          time = time_clean,
          event = event_clean
        )
        surv_fit <- survfit(Surv(time, event) ~ 1, data = surv_data)
        group_levels <- "Overall"
        n_groups <- 1
      }
      
      # Calculate summary statistics
      summary_stats <- summary(surv_fit)
      median_survival <- tryCatch({
        surv_median(surv_fit)
      }, error = function(e) {
        NULL
      })
      
      # Log-rank test if groups exist
      logrank_test <- NULL
      if (input$surv_group != "none" && n_groups > 1) {
        logrank_test <- tryCatch({
          survdiff(Surv(time, event) ~ group, data = surv_data)
        }, error = function(e) {
          NULL
        })
      }
      
      list(
        fit = surv_fit,
        data = surv_data,
        surv_obj = surv_obj,
        group_levels = group_levels,
        n_groups = n_groups,
        summary = summary_stats,
        median_survival = median_survival,
        logrank_test = logrank_test,
        success = TRUE
      )
      
    }, error = function(e) {
      return(list(
        fit = NULL,
        error = paste("Error in Kaplan-Meier analysis:", e$message),
        success = FALSE
      ))
    })
  })
  
  # Reactive for life table times
  life_table_times <- reactive({
    req(input$life_table_times)
    
    # Parse comma-separated time points
    times <- strsplit(input$life_table_times, ",")[[1]]
    times <- trimws(times)  # Remove whitespace
    times <- as.numeric(times)  # Convert to numeric
    
    # Remove any NA values and ensure times are positive
    times <- times[!is.na(times) & times >= 0]
    
    if (length(times) == 0) {
      # Default times if parsing fails
      return(c(0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0))
    }
    
    return(times)
  })
  
  # Life table results - CORRECTED VERSION
  life_table_results <- eventReactive(input$update_life_table, {
    req(surv_results(), input$life_table_times)
    
    results <- surv_results()
    
    if (!results$success) {
      return(list(
        table = data.frame(Error = "No survival analysis results available"),
        summary = "Please run the Kaplan-Meier analysis first.",
        success = FALSE
      ))
    }
    
    tryCatch({
      # Parse and validate time points
      times <- strsplit(input$life_table_times, ",")[[1]]
      times <- trimws(times)
      times <- as.numeric(times)
      times <- times[!is.na(times) & times >= 0]
      
      if (length(times) == 0) {
        return(list(
          table = data.frame(Error = "No valid time points provided"),
          summary = "Please enter valid numeric time points separated by commas.",
          success = FALSE
        ))
      }
      
      # Sort time points
      times <- sort(unique(times))
      
      # Get survival probabilities at specified times
      surv_summary <- summary(results$fit, times = times, extend = TRUE)
      
      # Create life table data frame
      if (results$n_groups == 1) {
        # Single group
        life_table <- data.frame(
          Time = surv_summary$time,
          Survival_Probability = round(surv_summary$surv, input$life_table_decimals),
          Standard_Error = round(surv_summary$std.err, input$life_table_decimals),
          Lower_CI = round(surv_summary$lower, input$life_table_decimals),
          Upper_CI = round(surv_summary$upper, input$life_table_decimals),
          Number_at_Risk = surv_summary$n.risk,
          Number_of_Events = surv_summary$n.event,
          Number_Censored = surv_summary$n.censor
        )
      } else {
        # Multiple groups
        life_table <- data.frame(
          Group = surv_summary$strata,
          Time = surv_summary$time,
          Survival_Probability = round(surv_summary$surv, input$life_table_decimals),
          Standard_Error = round(surv_summary$std.err, input$life_table_decimals),
          Lower_CI = round(surv_summary$lower, input$life_table_decimals),
          Upper_CI = round(surv_summary$upper, input$life_table_decimals),
          Number_at_Risk = surv_summary$n.risk,
          Number_of_Events = surv_summary$n.event,
          Number_Censored = surv_summary$n.censor
        )
        
        # Clean group names
        life_table$Group <- gsub(".*=", "", life_table$Group)
      }
      
      # Calculate additional statistics
      total_observations <- nrow(results$data)
      total_events <- sum(results$data$event)
      total_censored <- total_observations - total_events
      
      summary_text <- paste(
        "Life Table Summary:\n",
        "Total Observations: ", total_observations, "\n",
        "Total Events: ", total_events, "\n", 
        "Total Censored: ", total_censored, "\n",
        "Number of Time Points: ", length(times), "\n",
        "Time Range: ", min(times), " to ", max(times), "\n",
        if (results$n_groups > 1) paste("Number of Groups: ", results$n_groups),
        sep = ""
      )
      
      list(
        table = life_table,
        summary = summary_text,
        times = times,
        success = TRUE
      )
      
    }, error = function(e) {
      return(list(
        table = data.frame(Error = paste("Error generating life table:", e$message)),
        summary = "Error in life table generation",
        success = FALSE
      ))
    })
  })
  
  # Life table output - CORRECTED
  output$life_table_output <- renderDT({
    req(life_table_results())
    
    results <- life_table_results()
    
    if (!results$success) {
      datatable(
        results$table,
        options = list(dom = 't', pageLength = 10),
        rownames = FALSE,
        caption = "Life Table - Error"
      )
    } else {
      # Create a simple datatable without the problematic styling
      datatable(
        results$table,
        options = list(
          dom = 'Blfrtip',
          pageLength = 15,
          scrollX = TRUE,
          buttons = c('copy', 'csv', 'excel')
        ),
        rownames = FALSE,
        caption = paste("Life Table - Survival Probabilities at", length(results$times), "Specified Time Points"),
        extensions = 'Buttons'
      ) # Removed the problematic formatStyle section
    }
  })
  
  # Life table summary output - CORRECTED
  output$life_table_summary <- renderPrint({
    req(life_table_results())
    
    results <- life_table_results()
    
    if (!results$success) {
      cat("Life Table Summary - Error\n")
      cat("==========================\n")
      cat(results$table$Error, "\n")
    } else {
      cat(results$summary)
      cat("\n\nSpecified Time Points:\n")
      cat(paste(results$times, collapse = ", "), "\n")
    }
  })
  
  # Download handler for life table (Excel format) - CORRECTED
  output$download_life_table <- downloadHandler(
    filename = function() {
      paste("life_table_", Sys.Date(), ".xlsx", sep = "")
    },
    content = function(file) {
      req(life_table_results())
      
      results <- life_table_results()
      
      if (!results$success) {
        # Create error workbook
        wb <- createWorkbook()
        addWorksheet(wb, "Error")
        writeData(wb, "Error", data.frame(Error = results$table$Error))
        saveWorkbook(wb, file, overwrite = TRUE)
      } else {
        # Create workbook with life table data
        wb <- createWorkbook()
        
        # Main life table sheet
        addWorksheet(wb, "Life Table")
        writeData(wb, "Life Table", results$table)
        
        # Summary sheet
        addWorksheet(wb, "Summary")
        summary_df <- data.frame(
          Parameter = c("Analysis Date", "Total Observations", "Total Events", 
                        "Total Censored", "Number of Time Points", "Time Range", "Time Points"),
          Value = c(
            as.character(Sys.Date()),
            nrow(surv_results()$data),
            sum(surv_results()$data$event),
            nrow(surv_results()$data) - sum(surv_results()$data$event),
            length(results$times),
            paste(min(results$times), "to", max(results$times)),
            paste(results$times, collapse = ", ")
          )
        )
        writeData(wb, "Summary", summary_df)
        
        # Analysis parameters sheet
        addWorksheet(wb, "Parameters")
        params_df <- data.frame(
          Parameter = c("Time Variable", "Event Variable", "Grouping Variable", 
                        "Time Points", "Decimal Places"),
          Value = c(
            input$surv_time,
            input$surv_event,
            ifelse(input$surv_group != "none", input$surv_group, "None"),
            input$life_table_times,
            input$life_table_decimals
          )
        )
        writeData(wb, "Parameters", params_df)
        
        # Style the main table
        style <- createStyle(halign = "center", valign = "center")
        addStyle(wb, "Life Table", style, rows = 1:(nrow(results$table) + 1), 
                 cols = 1:ncol(results$table), gridExpand = TRUE)
        
        # Auto-adjust column widths
        setColWidths(wb, "Life Table", cols = 1:ncol(results$table), widths = "auto")
        setColWidths(wb, "Summary", cols = 1:2, widths = "auto")
        setColWidths(wb, "Parameters", cols = 1:2, widths = "auto")
        
        saveWorkbook(wb, file, overwrite = TRUE)
      }
    }
  )
  
  # Auto-update life table when survival results change
  observeEvent(surv_results(), {
    if (!is.null(surv_results()) && surv_results()$success) {
      # Trigger life table update when new survival results are available
      life_table_results()
    }
  })
  
  # Reactive for plot options to update plot without re-running analysis
  plot_options <- reactive({
    list(
      surv_ci = input$surv_ci,
      surv_risktable = input$surv_risktable,
      surv_median = input$surv_median,
      surv_labels = input$surv_labels,
      surv_color1 = input$surv_color1,
      surv_color2 = input$surv_color2,
      surv_color3 = input$surv_color3,
      surv_title = input$surv_title,
      surv_xlab = input$surv_xlab,
      surv_ylab = input$surv_ylab,
      surv_title_size = input$surv_title_size,
      surv_x_size = input$surv_x_size,
      surv_y_size = input$surv_y_size,
      surv_label_size = input$surv_label_size
    )
  })
  
  # Update the surv_plot output to use plot_options
  output$surv_plot <- renderPlot({
    req(surv_results(), plot_options())
    
    results <- surv_results()
    opts <- plot_options()
    
    if (!results$success) {
      # Create error plot
      p <- ggplot() +
        annotate("text", x = 0.5, y = 0.5, 
                 label = paste("Error:", results$error), 
                 size = 5, color = "red", hjust = 0.5, vjust = 0.5) +
        theme_void() +
        labs(title = "Kaplan-Meier Plot - Error")
      return(p)
    }
    
    tryCatch({
      # Create color palette based on number of groups
      if (results$n_groups == 1) {
        colors <- opts$surv_color1
      } else if (results$n_groups == 2) {
        colors <- c(opts$surv_color1, opts$surv_color2)
      } else {
        colors <- c(opts$surv_color1, opts$surv_color2, opts$surv_color3)
        # For more than 3 groups, use color brewer
        if (results$n_groups > 3) {
          colors <- RColorBrewer::brewer.pal(results$n_groups, "Set1")
        }
      }
      
      # Create the Kaplan-Meier plot using survminer
      p <- ggsurvplot(
        results$fit,
        data = results$data,
        conf.int = opts$surv_ci,
        risk.table = opts$surv_risktable,
        palette = colors,
        title = opts$surv_title,
        xlab = opts$surv_xlab,
        ylab = opts$surv_ylab,
        legend.title = ifelse(input$surv_group != "none", input$surv_group, ""),
        ggtheme = theme_minimal(),
        font.title = c(opts$surv_title_size, "bold"),
        font.x = c(opts$surv_x_size, "plain"),
        font.y = c(opts$surv_y_size, "plain"),
        font.tickslab = c(12, "plain"),
        risk.table.height = 0.25,
        surv.median.line = ifelse(opts$surv_median, "hv", "none"),
        censor = TRUE,
        size = 1.2  # Thicker lines for better visibility
      )
      
      # Customize the plot further with bigger and clearer elements
      p$plot <- p$plot +
        theme(
          plot.title = element_text(size = opts$surv_title_size, face = "bold", hjust = 0.5),
          axis.title.x = element_text(size = opts$surv_x_size, face = "bold"),
          axis.title.y = element_text(size = opts$surv_y_size, face = "bold"),
          axis.text.x = element_text(size = opts$surv_x_size - 2),
          axis.text.y = element_text(size = opts$surv_y_size - 2),
          legend.title = element_text(size = opts$surv_x_size, face = "bold"),
          legend.text = element_text(size = opts$surv_x_size - 2),
          legend.position = "right"
        )
      
      # Add survival probability labels if requested
      if (opts$surv_labels) {
        # Extract survival data for labeling
        surv_summary <- summary(results$fit)
        
        if (results$n_groups == 1) {
          # For single group, add labels at key time points
          key_times <- quantile(surv_summary$time, probs = seq(0.1, 0.9, 0.2))
          key_points <- surv_summary$surv[findInterval(key_times, surv_summary$time)]
          
          p$plot <- p$plot +
            geom_point(data = data.frame(time = key_times, surv = key_points), 
                       aes(x = time, y = surv), size = 3, color = "darkred") +
            geom_text(data = data.frame(time = key_times, surv = key_points),
                      aes(x = time, y = surv, 
                          label = paste0("S(t)=", round(surv, 2), "\nt=", round(time, 1))),
                      size = opts$surv_label_size, vjust = -0.5, hjust = -0.1, color = "darkblue")
        } else {
          # For multiple groups, add labels at median survival times
          if (!is.null(results$median_survival)) {
            median_data <- results$median_survival
            median_data$surv <- 0.5  # Median survival probability
            
            p$plot <- p$plot +
              geom_point(data = median_data, 
                         aes(x = median, y = surv, color = strata), 
                         size = 3, shape = 18) +
              geom_text(data = median_data,
                        aes(x = median, y = surv, color = strata,
                            label = paste0("Median: ", round(median, 1))),
                        size = opts$surv_label_size, vjust = -1, hjust = -0.1, 
                        show.legend = FALSE)
          }
        }
      }
      
      if (opts$surv_risktable) {
        p$table <- p$table +
          theme(
            axis.title.x = element_text(size = max(opts$surv_x_size - 2, 10), face = "bold"),
            axis.text.x = element_text(size = max(opts$surv_x_size - 2, 10)),
            plot.title = element_text(size = max(opts$surv_title_size - 2, 12), face = "bold")
          )
      }
      
      # Print the combined plot
      p
      
    }, error = function(e) {
      # Fallback error plot if ggsurvplot fails
      ggplot() +
        annotate("text", x = 0.5, y = 0.5, 
                 label = paste("Plotting Error:", e$message), 
                 size = 4, color = "red", hjust = 0.5, vjust = 0.5) +
        theme_void() +
        labs(title = "Kaplan-Meier Plot - Error")
    })
  })
  
  # Add summary statistics UI
  output$surv_stats <- renderUI({
    req(surv_results())
    
    results <- surv_results()
    
    if (!results$success) {
      return(
        div(
          class = "alert alert-danger",
          h4("Kaplan-Meier Analysis Error"),
          p(results$error)
        )
      )
    }
    
    div(
      class = "panel panel-default",
      div(
        class = "panel-heading",
        h4("Survival Analysis Summary", class = "panel-title")
      ),
      div(
        class = "panel-body",
        fluidRow(
          column(6,
                 tags$ul(
                   tags$li(tags$strong("Total Observations:"), nrow(results$data)),
                   tags$li(tags$strong("Events:"), sum(results$data$event)),
                   tags$li(tags$strong("Censored:"), sum(results$data$event == 0))
                 )
          ),
          column(6,
                 tags$ul(
                   tags$li(tags$strong("Groups:"), results$n_groups),
                   if (results$n_groups > 1 && !is.null(results$logrank_test)) {
                     tagList(
                       tags$li(tags$strong("Log-Rank p-value:"), 
                               round(1 - pchisq(results$logrank_test$chisq, length(results$logrank_test$n) - 1), 4))
                     )
                   }
                 )
          )
        )
      )
    )
  })
  
  output$surv_summary <- renderPrint({
    req(surv_results())
    
    results <- surv_results()
    
    if (!results$success) {
      cat("Kaplan-Meier Analysis Error:\n")
      cat(results$error, "\n")
      return()
    }
    
    cat("Kaplan-Meier Survival Analysis Summary\n")
    cat("======================================\n\n")
    
    # Print basic fit summary
    print(results$fit)
    
    if (!is.null(results$median_survival)) {
      cat("\nMedian Survival Time:\n")
      print(results$median_survival)
    }
    
    if (!is.null(results$logrank_test)) {
      cat("\nLog-Rank Test (Group Comparison):\n")
      print(results$logrank_test)
      cat("p-value:", round(1 - pchisq(results$logrank_test$chisq, length(results$logrank_test$n) - 1), 4), "\n")
    }
    
    # Additional summary statistics
    cat("\nSummary Statistics:\n")
    cat("- Number of observations:", nrow(results$data), "\n")
    cat("- Number of events:", sum(results$data$event), "\n")
    cat("- Censored:", sum(results$data$event == 0), "\n")
    
    if (results$n_groups > 1) {
      cat("- Number of groups:", results$n_groups, "\n")
      cat("- Group levels:", paste(results$group_levels, collapse = ", "), "\n")
    }
  })
  
  output$download_surv <- downloadHandler(
    filename = function() {
      paste("kaplan_meier_plot_", Sys.Date(), ".png", sep = "")
    },
    content = function(file) {
      req(surv_results(), plot_options())
      
      results <- surv_results()
      opts <- plot_options()
      
      if (!results$success) {
        # Create error plot
        p <- ggplot() +
          annotate("text", x = 0.5, y = 0.5, 
                   label = paste("Error:", results$error), 
                   size = 5, color = "red", hjust = 0.5, vjust = 0.5) +
          theme_void() +
          labs(title = "Kaplan-Meier Plot - Error")
        
        ggsave(file, plot = p, device = "png", width = 12, height = 10, dpi = 300)
      } else {
        # Create color palette for download
        if (results$n_groups == 1) {
          colors <- opts$surv_color1
        } else if (results$n_groups == 2) {
          colors <- c(opts$surv_color1, opts$surv_color2)
        } else {
          colors <- c(opts$surv_color1, opts$surv_color2, opts$surv_color3)
          if (results$n_groups > 3) {
            colors <- RColorBrewer::brewer.pal(results$n_groups, "Set1")
          }
        }
        
        # Create high-quality KM plot for download
        p_download <- ggsurvplot(
          results$fit,
          data = results$data,
          conf.int = opts$surv_ci,
          risk.table = opts$surv_risktable,
          palette = colors,
          title = opts$surv_title,
          xlab = opts$surv_xlab,
          ylab = opts$surv_ylab,
          legend.title = ifelse(input$surv_group != "none", input$surv_group, ""),
          ggtheme = theme_minimal(),
          font.title = c(18, "bold"),
          font.x = c(14, "bold"),
          font.y = c(14, "bold"),
          font.tickslab = c(12, "plain"),
          risk.table.height = 0.25,
          surv.median.line = ifelse(opts$surv_median, "hv", "none"),
          size = 1.5
        )
        
        # Add labels for download version as well
        if (opts$surv_labels) {
          surv_summary <- summary(results$fit)
          
          if (results$n_groups == 1) {
            key_times <- quantile(surv_summary$time, probs = seq(0.1, 0.9, 0.2))
            key_points <- surv_summary$surv[findInterval(key_times, surv_summary$time)]
            
            p_download$plot <- p_download$plot +
              geom_point(data = data.frame(time = key_times, surv = key_points), 
                         aes(x = time, y = surv), size = 4, color = "darkred") +
              geom_text(data = data.frame(time = key_times, surv = key_points),
                        aes(x = time, y = surv, 
                            label = paste0("S(t)=", round(surv, 2), "\nt=", round(time, 1))),
                        size = 5, vjust = -0.5, hjust = -0.1, color = "darkblue")
          }
        }
        
        # Customize the download plot theme
        p_download$plot <- p_download$plot +
          theme(
            plot.title = element_text(size = 18, face = "bold", hjust = 0.5),
            axis.title.x = element_text(size = 14, face = "bold"),
            axis.title.y = element_text(size = 14, face = "bold"),
            legend.title = element_text(size = 12, face = "bold"),
            legend.text = element_text(size = 11)
          )
        
        if (opts$surv_risktable) {
          p_download$table <- p_download$table +
            theme(
              axis.title.x = element_text(size = 12, face = "bold"),
              axis.text.x = element_text(size = 11),
              plot.title = element_text(size = 12, face = "bold")
            )
        }
        
        # Save the plot with high quality
        ggsave(file, plot = print(p_download), device = "png", 
               width = 14, height = 10, dpi = 300, bg = "white")
      }
    }
  )
  
  # Add data download
  output$download_surv_data <- downloadHandler(
    filename = function() {
      paste("survival_data_", Sys.Date(), ".csv", sep = "")
    },
    content = function(file) {
      req(surv_results())
      
      results <- surv_results()
      
      if (!results$success) {
        # Create error data frame
        write.csv(data.frame(Error = results$error), file, row.names = FALSE)
      } else {
        # Write survival data to CSV
        surv_data <- results$data %>%
          select(Time = time, Event = event)
        
        if (input$surv_group != "none") {
          surv_data$Group <- results$data$group
        }
        
        write.csv(surv_data, file, row.names = FALSE)
      }
    }
  )
  
  # Cox Regression reference UI
  output$cox_ref_ui <- renderUI({
    req(input$cox_vars, processed_data())
    df <- processed_data()
    
    ref_inputs <- tagList()
    
    for(var in input$cox_vars) {
      if(is.factor(df[[var]]) || is.character(df[[var]])) {
        levels <- if(is.factor(df[[var]])) levels(df[[var]]) else unique(na.omit(df[[var]]))
        ref_inputs <- tagAppendChild(
          ref_inputs,
          selectInput(paste0("cox_ref_", var), 
                      paste("Reference level for", var), 
                      choices = levels)
        )
      }
    }
    
    ref_inputs
  })
  
  # Cox Regression with action button
  cox_results <- eventReactive(input$run_cox, {
    req(input$cox_time, input$cox_event, input$cox_vars, processed_data())
    df <- processed_data()
    
    # Set reference levels
    for(var in input$cox_vars) {
      ref_input <- input[[paste0("cox_ref_", var)]]
      if(!is.null(ref_input) && (is.factor(df[[var]]) || is.character(df[[var]]))) {
        df[[var]] <- factor(df[[var]])
        df[[var]] <- relevel(df[[var]], ref = ref_input)
      }
    }
    
    # Create survival object
    surv_obj <- create_surv_obj(input$cox_time, input$cox_event, df)
    
    # Create formula
    formula <- as.formula(paste("surv_obj ~", paste(input$cox_vars, collapse = "+")))
    
    # Fit Cox model
    model <- coxph(formula, data = df)
    
    list(model = model, summary = summary(model))
  })
  
  output$cox_summary <- renderPrint({
    req(cox_results())
    print(cox_results()$summary)
  })
  
  output$cox_table <- renderDT({
    req(cox_results())
    cox_table <- broom::tidy(cox_results()$model, conf.int = TRUE, conf.level = input$cox_conf, exponentiate = TRUE)
    datatable(cox_table, options = list(pageLength = 10))
  })
  
  output$download_cox <- downloadHandler(
    filename = "cox_regression.docx",
    content = function(file) {
      cox_table <- broom::tidy(cox_results()$model, conf.int = TRUE, conf.level = input$cox_conf, exponentiate = TRUE)
      
      ft <- flextable(cox_table) %>%
        set_caption("Cox Regression Results") %>%
        theme_zebra()
      
      doc <- read_docx() %>% 
        body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # Poisson Regression reference UI
  output$poisson_ref_ui <- renderUI({
    req(input$poisson_vars, processed_data())
    df <- processed_data()
    
    ref_inputs <- tagList()
    
    for(var in input$poisson_vars) {
      if(is.factor(df[[var]]) || is.character(df[[var]])) {
        levels <- if(is.factor(df[[var]])) levels(df[[var]]) else unique(na.omit(df[[var]]))
        ref_inputs <- tagAppendChild(
          ref_inputs,
          selectInput(paste0("poisson_ref_", var), 
                      paste("Reference level for", var), 
                      choices = levels)
        )
      }
    }
    
    ref_inputs
  })
  
  # Poisson Regression with action button - CORRECTED VERSION
  poisson_results <- eventReactive(input$run_poisson, {
    req(input$poisson_outcome, input$poisson_vars, processed_data())
    
    tryCatch({
      df <- processed_data()
      
      # Validate variables exist
      if (!input$poisson_outcome %in% names(df)) {
        return(list(
          error = "Outcome variable not found in dataset",
          success = FALSE
        ))
      }
      
      missing_vars <- setdiff(input$poisson_vars, names(df))
      if (length(missing_vars) > 0) {
        return(list(
          error = paste("Predictor variables not found:", paste(missing_vars, collapse = ", ")),
          success = FALSE
        ))
      }
      
      # Set reference levels
      for(var in input$poisson_vars) {
        ref_input <- input[[paste0("poisson_ref_", var)]]
        if(!is.null(ref_input) && (is.factor(df[[var]]) || is.character(df[[var]]))) {
          df[[var]] <- factor(df[[var]])
          df[[var]] <- relevel(df[[var]], ref = ref_input)
        }
      }
      
      # Create formula
      formula <- as.formula(paste(input$poisson_outcome, "~", paste(input$poisson_vars, collapse = "+")))
      
      # Add offset if specified
      if(input$poisson_offset && !is.null(input$poisson_offset_var)) {
        if (!input$poisson_offset_var %in% names(df)) {
          return(list(
            error = "Offset variable not found in dataset",
            success = FALSE
          ))
        }
        formula <- as.formula(paste(input$poisson_outcome, "~", paste(input$poisson_vars, collapse = "+"), 
                                    "+ offset(log(", input$poisson_offset_var, "))"))
      }
      
      # Remove rows with missing values
      df_complete <- df[complete.cases(df[, c(input$poisson_outcome, input$poisson_vars)]), ]
      
      if (nrow(df_complete) == 0) {
        return(list(
          error = "No complete cases after removing missing values",
          success = FALSE
        ))
      }
      
      # Check if outcome is count data (non-negative integers)
      outcome_data <- df_complete[[input$poisson_outcome]]
      if (any(outcome_data < 0, na.rm = TRUE)) {
        return(list(
          error = "Outcome variable contains negative values. Poisson regression requires non-negative counts.",
          success = FALSE
        ))
      }
      
      # Fit Poisson model
      model <- glm(formula, family = poisson(link = "log"), data = df_complete)
      
      # Check for model convergence
      if (!model$converged) {
        warning("Poisson model did not converge properly")
      }
      
      # Calculate confidence intervals
      conf_level <- input$poisson_conf
      z_value <- qnorm(1 - (1 - conf_level) / 2)
      
      list(
        model = model,
        summary = summary(model),
        data = df_complete,
        success = TRUE
      )
      
    }, error = function(e) {
      return(list(
        error = paste("Error in Poisson regression:", e$message),
        success = FALSE
      ))
    })
  })
  
  output$poisson_summary <- renderPrint({
    req(poisson_results())
    
    results <- poisson_results()
    
    if (!results$success) {
      cat("POISSON REGRESSION ANALYSIS - ERROR\n")
      cat("===================================\n")
      cat(results$error, "\n")
      return()
    }
    
    cat("POISSON REGRESSION ANALYSIS\n")
    cat("===========================\n\n")
    
    # Model information
    cat("MODEL INFORMATION:\n")
    cat("- Outcome variable:", input$poisson_outcome, "\n")
    cat("- Predictor variables:", paste(input$poisson_vars, collapse = ", "), "\n")
    if (input$poisson_offset && !is.null(input$poisson_offset_var)) {
      cat("- Offset variable:", input$poisson_offset_var, "\n")
    }
    cat("- Observations:", nrow(results$data), "\n")
    cat("- Confidence level:", input$poisson_conf * 100, "%\n")
    cat("\n")
    
    # Model summary
    print(results$summary)
    
    # Additional model diagnostics
    cat("\nMODEL DIAGNOSTICS:\n")
    cat("- Null deviance:", round(results$model$null.deviance, 3), "\n")
    cat("- Residual deviance:", round(results$model$deviance, 3), "\n")
    cat("- AIC:", round(AIC(results$model), 3), "\n")
    cat("- Dispersion parameter:", round(results$summary$dispersion, 3), "\n")
    
    # Check for overdispersion
    if (results$summary$dispersion > 1.5) {
      cat("⚠ Warning: Possible overdispersion detected. Consider using negative binomial regression.\n")
    }
  })
  
  output$poisson_table <- renderDT({
    req(poisson_results())
    
    results <- poisson_results()
    
    if (!results$success) {
      return(datatable(
        data.frame(Error = results$error),
        options = list(dom = 't'),
        rownames = FALSE,
        caption = "Poisson Regression - Error"
      ))
    }
    
    # Create results table with rate ratios and confidence intervals
    poisson_table <- broom::tidy(results$model, conf.int = TRUE, conf.level = input$poisson_conf, exponentiate = TRUE)
    
    # Add log-scale estimates
    poisson_table_log <- broom::tidy(results$model, conf.int = TRUE, conf.level = input$poisson_conf, exponentiate = FALSE)
    poisson_table$log_estimate <- poisson_table_log$estimate
    poisson_table$log_std.error <- poisson_table_log$std.error
    poisson_table$log_conf.low <- poisson_table_log$conf.low
    poisson_table$log_conf.high <- poisson_table_log$conf.high
    
    datatable(
      poisson_table,
      options = list(
        pageLength = 15,
        dom = 'Blfrtip',
        scrollX = TRUE
      ),
      rownames = FALSE,
      caption = "Poisson Regression Results (Rate Ratios and 95% Confidence Intervals)"
    ) %>%
      formatRound(columns = c('estimate', 'std.error', 'statistic', 'p.value', 'conf.low', 'conf.high',
                              'log_estimate', 'log_std.error', 'log_conf.low', 'log_conf.high'), 
                  digits = 4)
  })
  
  output$download_poisson <- downloadHandler(
    filename = function() {
      paste("poisson_regression_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(poisson_results())
      
      results <- poisson_results()
      
      if (!results$success) {
        # Create error document
        ft <- flextable(data.frame(Error = results$error)) %>%
          set_caption("Poisson Regression Analysis - Error") %>%
          theme_zebra()
        
        doc <- read_docx() %>% 
          body_add_flextable(ft)
      } else {
        # Create comprehensive results table
        poisson_table <- broom::tidy(results$model, conf.int = TRUE, conf.level = input$poisson_conf, exponentiate = TRUE)
        
        # Add log-scale estimates
        poisson_table_log <- broom::tidy(results$model, conf.int = TRUE, conf.level = input$poisson_conf, exponentiate = FALSE)
        poisson_table$log_estimate <- poisson_table_log$estimate
        poisson_table$log_std.error <- poisson_table_log$std.error
        poisson_table$log_conf.low <- poisson_table_log$conf.low
        poisson_table$log_conf.high <- poisson_table_log$conf.high
        
        # Create report
        doc <- read_docx()
        
        # Title and overview
        doc <- doc %>%
          body_add_par("POISSON REGRESSION ANALYSIS REPORT", style = "heading 1") %>%
          body_add_par(paste("Outcome Variable:", input$poisson_outcome), style = "Normal") %>%
          body_add_par(paste("Predictor Variables:", paste(input$poisson_vars, collapse = ", ")), style = "Normal")
        
        if (input$poisson_offset && !is.null(input$poisson_offset_var)) {
          doc <- doc %>% body_add_par(paste("Offset Variable:", input$poisson_offset_var), style = "Normal")
        }
        
        doc <- doc %>%
          body_add_par(paste("Confidence Level:", input$poisson_conf * 100, "%"), style = "Normal") %>%
          body_add_par(paste("Observations:", nrow(results$data)), style = "Normal") %>%
          body_add_par("", style = "Normal")
        
        # Model coefficients
        doc <- doc %>%
          body_add_par("REGRESSION COEFFICIENTS (RATE RATIOS)", style = "heading 2")
        
        ft <- flextable(poisson_table) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft)
        
        # Model diagnostics
        doc <- doc %>%
          body_add_par("MODEL DIAGNOSTICS", style = "heading 2")
        
        diagnostics <- data.frame(
          Statistic = c("Null Deviance", "Residual Deviance", "AIC", "Dispersion Parameter"),
          Value = c(
            round(results$model$null.deviance, 3),
            round(results$model$deviance, 3),
            round(AIC(results$model), 3),
            round(results$summary$dispersion, 3)
          )
        )
        
        ft_diag <- flextable(diagnostics) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_diag)
      }
      
      print(doc, target = file)
    }
  )
  # Linear Regression reference UI
  output$linear_ref_ui <- renderUI({
    req(input$linear_vars, processed_data())
    df <- processed_data()
    
    ref_inputs <- tagList()
    
    for(var in input$linear_vars) {
      if(is.factor(df[[var]]) || is.character(df[[var]])) {
        levels <- if(is.factor(df[[var]])) levels(df[[var]]) else unique(na.omit(df[[var]]))
        ref_inputs <- tagAppendChild(
          ref_inputs,
          selectInput(paste0("linear_ref_", var), 
                      paste("Reference level for", var), 
                      choices = levels)
        )
      }
    }
    
    ref_inputs
  })
  
  # Linear Regression with action button
  linear_results <- eventReactive(input$run_linear, {
    req(input$linear_outcome, input$linear_vars, processed_data())
    df <- processed_data()
    
    # Set reference levels
    for(var in input$linear_vars) {
      ref_input <- input[[paste0("linear_ref_", var)]]
      if(!is.null(ref_input) && (is.factor(df[[var]]) || is.character(df[[var]]))) {
        df[[var]] <- factor(df[[var]])
        df[[var]] <- relevel(df[[var]], ref = ref_input)
      }
    }
    
    # Create formula
    formula <- as.formula(paste(input$linear_outcome, "~", paste(input$linear_vars, collapse = "+")))
    
    # Fit linear model
    model <- lm(formula, data = df)
    
    list(model = model, summary = summary(model))
  })
  
  output$linear_summary <- renderPrint({
    req(linear_results())
    print(linear_results()$summary)
    
    if(input$linear_std) {
      cat("\nStandardized Coefficients:\n")
      std_coef <- lm.beta(linear_results()$model)
      print(std_coef)
    }
  })
  
  output$linear_table <- renderDT({
    req(linear_results())
    linear_table <- broom::tidy(linear_results()$model, conf.int = TRUE)
    
    if(input$linear_std) {
      std_coef <- lm.beta(linear_results()$model)$standardized.coefficients
      linear_table$std_estimate <- std_coef
    }
    
    datatable(linear_table, options = list(pageLength = 10))
  })
  
  output$download_linear <- downloadHandler(
    filename = "linear_regression.docx",
    content = function(file) {
      linear_table <- broom::tidy(linear_results()$model, conf.int = TRUE)
      
      ft <- flextable(linear_table) %>%
        set_caption("Linear Regression Results") %>%
        theme_zebra()
      
      doc <- read_docx() %>% 
        body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # Server logic for logistic regression - CORRECTED VERSION (3 decimal places)
  
  # UI for outcome variable handling
  output$logistic_outcome_ui <- renderUI({
    req(input$logistic_outcome, processed_data())
    df <- processed_data()
    
    if (!input$logistic_outcome %in% names(df)) return(NULL)
    
    outcome_var <- df[[input$logistic_outcome]]
    
    # Check if outcome is numeric binary (0/1) or needs conversion
    if (is.numeric(outcome_var)) {
      unique_vals <- unique(na.omit(outcome_var))
      if (length(unique_vals) == 2 && all(unique_vals %in% c(0, 1))) {
        # Already numeric binary (0/1)
        return(
          div(
            class = "alert alert-success",
            h5("Outcome Variable Status"),
            p("✓ Outcome variable is numeric binary (0/1)"),
            p(paste("Levels: 0 =", names(which.min(table(outcome_var))), 
                    ", 1 =", names(which.max(table(outcome_var)))))
          )
        )
      }
    }
    
    # If not numeric binary, provide reference level selection
    tagList(
      div(
        class = "alert alert-info",
        h5("Outcome Variable Configuration"),
        p("Please specify the reference category for the outcome variable:")
      ),
      selectInput("logistic_outcome_ref", 
                  "Reference Category (Baseline):",
                  choices = if (is.factor(outcome_var)) {
                    levels(outcome_var)
                  } else {
                    unique(na.omit(outcome_var))
                  }),
      helpText("The reference category will be coded as 0 in the analysis.")
    )
  })
  
  # UI for predictor variable reference levels
  output$logistic_ref_ui <- renderUI({
    req(input$logistic_vars, processed_data())
    df <- processed_data()
    
    ref_inputs <- tagList(
      h5("Reference Levels for Predictor Variables"),
      p("Specify reference categories for categorical predictors:")
    )
    
    for(var in input$logistic_vars) {
      if (var %in% names(df)) {
        var_data <- df[[var]]
        
        # Only show reference selection for categorical variables with reasonable number of levels
        if ((is.factor(var_data) || is.character(var_data)) && length(unique(na.omit(var_data))) <= 20) {
          
          levels <- if (is.factor(var_data)) {
            levels(var_data)
          } else {
            sort(unique(na.omit(var_data)))
          }
          
          ref_inputs <- tagAppendChild(
            ref_inputs,
            selectInput(paste0("logistic_ref_", make.names(var)), 
                        paste("Reference for", var), 
                        choices = levels,
                        selected = levels[1])
          )
        } else if ((is.factor(var_data) || is.character(var_data)) && length(unique(na.omit(var_data))) > 20) {
          ref_inputs <- tagAppendChild(
            ref_inputs,
            div(
              class = "alert alert-warning",
              p(paste("⚠ Variable", var, "has", length(unique(na.omit(var_data))), 
                      "levels. Consider recoding or using as continuous variable."))
            )
          )
        }
      }
    }
    
    if (length(ref_inputs) == 2) { # Only has header and no variables
      ref_inputs <- tagList(
        h5("Reference Levels for Predictor Variables"),
        p("No categorical predictors selected or all predictors are numeric.")
      )
    }
    
    ref_inputs
  })
  
  # Logistic Regression Analysis - CORRECTED VERSION (3 decimal places)
  logistic_results <- eventReactive(input$run_logistic, {
    req(input$logistic_outcome, input$logistic_vars, processed_data())
    
    tryCatch({
      df <- processed_data()
      
      # Validate variables exist
      if (!input$logistic_outcome %in% names(df)) {
        stop("Outcome variable not found in dataset")
      }
      
      missing_vars <- setdiff(input$logistic_vars, names(df))
      if (length(missing_vars) > 0) {
        stop(paste("Predictor variables not found:", paste(missing_vars, collapse = ", ")))
      }
      
      # Prepare outcome variable
      outcome_var <- df[[input$logistic_outcome]]
      
      # Convert outcome to binary numeric (0/1)
      if (is.numeric(outcome_var)) {
        unique_vals <- unique(na.omit(outcome_var))
        if (length(unique_vals) == 2 && all(unique_vals %in% c(0, 1))) {
          # Already numeric binary
          df$outcome_binary <- outcome_var
          outcome_levels <- c("0" = "Reference", "1" = "Comparison")
        } else {
          # Numeric but not binary - convert to binary based on median
          cutoff <- median(outcome_var, na.rm = TRUE)
          df$outcome_binary <- as.numeric(outcome_var > cutoff)
          outcome_levels <- c("0" = paste("≤", round(cutoff, 3)), "1" = paste(">", round(cutoff, 3)))
          warning("Numeric outcome converted to binary using median cutoff")
        }
      } else {
        # Categorical outcome - use user-specified reference
        if (is.null(input$logistic_outcome_ref)) {
          stop("Please specify reference category for outcome variable")
        }
        
        outcome_factor <- factor(outcome_var)
        ref_level <- input$logistic_outcome_ref
        
        # Set reference level
        if (ref_level %in% levels(outcome_factor)) {
          outcome_factor <- relevel(outcome_factor, ref = ref_level)
        }
        
        # Convert to binary numeric
        df$outcome_binary <- as.numeric(outcome_factor) - 1
        outcome_levels <- setNames(levels(outcome_factor), c("0", "1"))
      }
      
      # Remove rows with missing outcome
      df <- df[!is.na(df$outcome_binary), ]
      
      # Check if we have both outcome levels
      if (length(unique(df$outcome_binary)) != 2) {
        stop("Outcome variable must have exactly two categories after processing")
      }
      
      # Prepare predictor variables
      for(var in input$logistic_vars) {
        if (var %in% names(df)) {
          var_data <- df[[var]]
          
          # Handle categorical predictors
          if (is.factor(var_data) || is.character(var_data)) {
            # Get user-specified reference level
            ref_input_name <- paste0("logistic_ref_", make.names(var))
            ref_level <- input[[ref_input_name]]
            
            if (!is.null(ref_level) && ref_level %in% unique(na.omit(var_data))) {
              df[[var]] <- factor(var_data)
              df[[var]] <- relevel(df[[var]], ref = ref_level)
            } else {
              # Use default reference (first level)
              df[[var]] <- factor(var_data)
            }
          }
          
          # Handle numeric predictors - ensure they're numeric
          if (is.character(var_data) && !all(is.na(suppressWarnings(as.numeric(na.omit(var_data)))))) {
            df[[var]] <- as.numeric(var_data)
          }
        }
      }
      
      # Remove rows with missing values in predictors
      complete_cases <- complete.cases(df[, input$logistic_vars, drop = FALSE])
      df <- df[complete_cases, ]
      
      if (nrow(df) == 0) {
        stop("No complete cases after removing missing values")
      }
      
      # Check for complete separation or other issues
      if (any(table(df$outcome_binary) < 2)) {
        stop("Insufficient cases in one of the outcome categories")
      }
      
      # Create formula
      formula <- as.formula(paste("outcome_binary ~", paste(input$logistic_vars, collapse = "+")))
      
      # Fit logistic regression model
      model <- glm(formula, family = binomial(link = "logit"), data = df)
      
      # Check for model convergence
      if (!model$converged) {
        warning("Model did not converge properly")
      }
      
      # Calculate model performance metrics
      predictions <- predict(model, type = "response")
      predicted_classes <- as.numeric(predictions > 0.5)
      accuracy <- mean(predicted_classes == df$outcome_binary, na.rm = TRUE)
      
      # Calculate AUC if possible
      auc_value <- tryCatch({
        roc_obj <- pROC::roc(df$outcome_binary, predictions, quiet = TRUE)
        as.numeric(pROC::auc(roc_obj))
      }, error = function(e) NA)
      
      list(
        model = model,
        summary = summary(model),
        data = df,
        outcome_levels = outcome_levels,
        performance = list(
          accuracy = accuracy,
          auc = auc_value,
          n_observations = nrow(df),
          n_events = sum(df$outcome_binary),
          n_nonevents = sum(1 - df$outcome_binary)
        ),
        success = TRUE
      )
      
    }, error = function(e) {
      return(list(
        error = paste("Error in logistic regression:", e$message),
        success = FALSE
      ))
    })
  })
  
  # Output for logistic regression results - CORRECTED (3 decimal places)
  output$logistic_summary <- renderPrint({
    req(logistic_results())
    
    results <- logistic_results()
    
    if (!results$success) {
      cat("LOGISTIC REGRESSION ANALYSIS - ERROR\n")
      cat("====================================\n")
      cat(results$error, "\n")
      return()
    }
    
    cat("LOGISTIC REGRESSION ANALYSIS\n")
    cat("============================\n\n")
    
    # Outcome variable information
    cat("OUTCOME VARIABLE:\n")
    cat("- Variable:", input$logistic_outcome, "\n")
    cat("- Coding: ", paste(names(results$outcome_levels), "=", results$outcome_levels, collapse = ", "), "\n")
    cat("\n")
    
    # Model performance
    cat("MODEL PERFORMANCE:\n")
    cat("- Observations:", results$performance$n_observations, "\n")
    cat("- Events:", results$performance$n_events, "\n")
    cat("- Non-events:", results$performance$n_nonevents, "\n")
    cat("- Accuracy:", round(results$performance$accuracy, 3), "\n")
    if (!is.na(results$performance$auc)) {
      cat("- AUC:", round(results$performance$auc, 3), "\n")
    }
    cat("\n")
    
    # Model summary with formatted coefficients
    cat("MODEL SUMMARY:\n")
    model_summary <- results$summary
    coef_matrix <- model_summary$coefficients
    
    # Format coefficients to 3 decimal places
    formatted_coef <- matrix(
      sapply(coef_matrix, function(x) {
        if (is.na(x)) return("NA")
        if (abs(x) < 0.001 & abs(x) > 0) {
          formatC(x, format = "e", digits = 3)
        } else {
          format(round(x, 3), nsmall = 3, scientific = FALSE)
        }
      }),
      nrow = nrow(coef_matrix),
      ncol = ncol(coef_matrix)
    )
    
    rownames(formatted_coef) <- rownames(coef_matrix)
    colnames(formatted_coef) <- colnames(coef_matrix)
    
    print(formatted_coef, quote = FALSE)
    
    # Odds ratios if requested - formatted to 3 decimal places
    if (input$logistic_or) {
      cat("\nODDS RATIOS:\n")
      coef_table <- coef(model_summary)
      odds_ratios <- exp(coef_table[, 1])
      or_ci_lower <- exp(coef_table[, 1] - 1.96 * coef_table[, 2])
      or_ci_upper <- exp(coef_table[, 1] + 1.96 * coef_table[, 2])
      
      # Format odds ratios to 3 decimal places
      or_table <- data.frame(
        Variable = rownames(coef_table),
        OddsRatio = format(round(odds_ratios, 3), nsmall = 3, scientific = FALSE),
        LowerCI = format(round(or_ci_lower, 3), nsmall = 3, scientific = FALSE),
        UpperCI = format(round(or_ci_upper, 3), nsmall = 3, scientific = FALSE),
        stringsAsFactors = FALSE
      )
      print(or_table, row.names = FALSE)
    }
    
    # Model fit statistics
    cat("\nMODEL FIT STATISTICS:\n")
    cat("- Null deviance:", round(results$model$null.deviance, 3), "\n")
    cat("- Residual deviance:", round(results$model$deviance, 3), "\n")
    cat("- AIC:", round(AIC(results$model), 3), "\n")
  })
  
  output$logistic_table <- renderDT({
    req(logistic_results())
    
    results <- logistic_results()
    
    if (!results$success) {
      return(datatable(
        data.frame(Error = results$error),
        options = list(dom = 't'),
        rownames = FALSE
      ))
    }
    
    # Create comprehensive results table with 3 decimal places
    if (input$logistic_or) {
      logistic_table <- broom::tidy(results$model, conf.int = TRUE, exponentiate = TRUE)
    } else {
      logistic_table <- broom::tidy(results$model, conf.int = TRUE)
    }
    
    # Format all numeric columns to 3 decimal places
    numeric_cols <- c("estimate", "std.error", "statistic", "p.value", "conf.low", "conf.high")
    for(col in numeric_cols) {
      if(col %in% names(logistic_table)) {
        logistic_table[[col]] <- format(round(logistic_table[[col]], 3), nsmall = 3, scientific = FALSE)
      }
    }
    
    datatable(
      logistic_table,
      options = list(
        pageLength = 15,
        dom = 'Blfrtip',
        scrollX = TRUE
      ),
      rownames = FALSE,
      caption = "Logistic Regression Coefficients (values rounded to 3 decimal places)"
    )
  })
  
  output$download_logistic <- downloadHandler(
    filename = function() {
      paste("logistic_regression_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(logistic_results())
      
      results <- logistic_results()
      
      if (!results$success) {
        # Create error document
        ft <- flextable(data.frame(Error = results$error)) %>%
          set_caption("Logistic Regression Analysis - Error") %>%
          theme_zebra()
        
        doc <- read_docx() %>% 
          body_add_flextable(ft)
      } else {
        # Create comprehensive report
        doc <- read_docx()
        
        # Title and overview
        doc <- doc %>%
          body_add_par("LOGISTIC REGRESSION ANALYSIS REPORT", style = "heading 1") %>%
          body_add_par(paste("Outcome Variable:", input$logistic_outcome), style = "Normal") %>%
          body_add_par(paste("Predictor Variables:", paste(input$logistic_vars, collapse = ", ")), style = "Normal") %>%
          body_add_par("", style = "Normal")
        
        # Outcome coding information
        doc <- doc %>%
          body_add_par("OUTCOME VARIABLE CODING", style = "heading 2")
        
        outcome_coding <- data.frame(
          Code = names(results$outcome_levels),
          Meaning = results$outcome_levels
        )
        ft_outcome <- flextable(outcome_coding) %>%
          theme_zebra() %>%
          autofit()
        doc <- doc %>% body_add_flextable(ft_outcome)
        
        # Model performance
        doc <- doc %>%
          body_add_par("MODEL PERFORMANCE", style = "heading 2")
        
        performance_df <- data.frame(
          Metric = c("Observations", "Events", "Non-events", "Accuracy", "AUC"),
          Value = c(
            results$performance$n_observations,
            results$performance$n_events,
            results$performance$n_nonevents,
            format(round(results$performance$accuracy, 3), nsmall = 3),
            ifelse(!is.na(results$performance$auc), format(round(results$performance$auc, 3), nsmall = 3), "N/A")
          )
        )
        ft_performance <- flextable(performance_df) %>%
          theme_zebra() %>%
          autofit()
        doc <- doc %>% body_add_flextable(ft_performance)
        
        # Coefficients table
        doc <- doc %>%
          body_add_par("REGRESSION COEFFICIENTS", style = "heading 2")
        
        if (input$logistic_or) {
          coef_table <- broom::tidy(results$model, conf.int = TRUE, exponentiate = TRUE)
        } else {
          coef_table <- broom::tidy(results$model, conf.int = TRUE)
        }
        
        # Format coefficients to 3 decimal places
        numeric_cols <- c("estimate", "std.error", "statistic", "p.value", "conf.low", "conf.high")
        for(col in numeric_cols) {
          if(col %in% names(coef_table)) {
            coef_table[[col]] <- format(round(coef_table[[col]], 3), nsmall = 3, scientific = FALSE)
          }
        }
        
        ft_coef <- flextable(coef_table) %>%
          theme_zebra() %>%
          autofit()
        doc <- doc %>% body_add_flextable(ft_coef)
        
        # Model fit statistics
        doc <- doc %>%
          body_add_par("MODEL FIT STATISTICS", style = "heading 2")
        
        fit_stats <- data.frame(
          Statistic = c("Null Deviance", "Residual Deviance", "AIC"),
          Value = c(
            format(round(results$model$null.deviance, 3), nsmall = 3),
            format(round(results$model$deviance, 3), nsmall = 3),
            format(round(AIC(results$model), 3), nsmall = 3)
          )
        )
        ft_fit <- flextable(fit_stats) %>%
          theme_zebra() %>%
          autofit()
        doc <- doc %>% body_add_flextable(ft_fit)
      }
      
      print(doc, target = file)
    }
  )
  
  # Chi-square test with action button
  chisq_results <- eventReactive(input$run_chisq, {
    req(input$chisq_var1, input$chisq_var2, processed_data())
    df <- processed_data()
    
    # Create contingency table
    tab <- table(df[[input$chisq_var1]], df[[input$chisq_var2]])
    
    # Perform chi-square test
    chisq_test <- chisq.test(tab)
    
    # Additional tests if requested
    fisher_test <- NULL
    rr_test <- NULL
    
    if(input$chisq_fisher) {
      fisher_test <- fisher.test(tab)
    }
    
    if(input$chisq_rr) {
      rr_test <- riskratio(tab)
    }
    
    # Create cross-tabulation with percentages
    cross_tab <- tab %>%
      as.data.frame.matrix() %>%
      rownames_to_column("Category") %>%
      mutate(Total = rowSums(across(where(is.numeric))))
    
    list(
      test = chisq_test,
      table = tab,
      cross_tab = cross_tab,
      fisher = fisher_test,
      rr = rr_test
    )
  })
  
  output$chisq_results <- renderPrint({
    req(chisq_results())
    
    cat("Chi-square Test Results:\n")
    print(chisq_results()$test)
    
    if(input$chisq_expected) {
      cat("\nExpected Counts:\n")
      print(chisq_results()$test$expected)
    }
    
    if(input$chisq_residuals) {
      cat("\nPearson Residuals:\n")
      print(chisq_results()$test$residuals)
    }
    
    if(input$chisq_fisher && !is.null(chisq_results()$fisher)) {
      cat("\nFisher's Exact Test:\n")
      print(chisq_results()$fisher)
    }
    
    if(input$chisq_rr && !is.null(chisq_results()$rr)) {
      cat("\nRelative Risk:\n")
      print(chisq_results()$rr)
    }
  })
  
  output$chisq_table <- renderDT({
    req(chisq_results())
    datatable(chisq_results()$cross_tab, 
              caption = "Cross-tabulation",
              options = list(pageLength = 10))
  })
  
  output$chisq_plot <- renderPlot({
    req(chisq_results(), input$chisq_var1, input$chisq_var2)
    df <- processed_data()
    
    # Create stacked bar plot
    p <- ggplot(df, aes(x = !!sym(input$chisq_var1), fill = !!sym(input$chisq_var2))) +
      geom_bar(position = "fill") +
      labs(title = paste("Stacked Bar Plot:", input$chisq_var1, "vs", input$chisq_var2),
           x = input$chisq_var1,
           y = "Proportion",
           fill = input$chisq_var2) +
      theme_minimal() +
      theme(axis.text.x = element_text(angle = 45, hjust = 1))
    
    print(p)
  })
  
  # CORRECTED Chi-square Test Download Handler
  output$download_chisq <- downloadHandler(
    filename = function() {
      paste("chi_square_test_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(chisq_results())
      
      results <- chisq_results()
      
      # Create comprehensive report
      doc <- read_docx()
      
      # Title and basic information
      doc <- doc %>%
        body_add_par("CHI-SQUARE TEST ANALYSIS REPORT", style = "heading 1") %>%
        body_add_par(paste("Variable 1:", input$chisq_var1), style = "Normal") %>%
        body_add_par(paste("Variable 2:", input$chisq_var2), style = "Normal") %>%
        body_add_par("", style = "Normal")
      
      # Contingency table
      doc <- doc %>%
        body_add_par("CONTINGENCY TABLE", style = "heading 2")
      
      # Convert table to data frame with proper formatting
      contingency_df <- as.data.frame.matrix(results$table)
      contingency_df <- cbind(Variable = rownames(contingency_df), contingency_df)
      
      ft_contingency <- flextable(contingency_df) %>%
        theme_zebra() %>%
        autofit()
      
      doc <- doc %>% body_add_flextable(ft_contingency)
      
      # Chi-square test results
      doc <- doc %>%
        body_add_par("CHI-SQUARE TEST RESULTS", style = "heading 2")
      
      chi_summary <- data.frame(
        Statistic = c("Chi-square Statistic", "Degrees of Freedom", "P-value", "Significance"),
        Value = c(
          round(results$test$statistic, 4),
          results$test$parameter,
          round(results$test$p.value, 4),
          ifelse(results$test$p.value < 0.05, "Statistically Significant", "Not Statistically Significant")
        )
      )
      
      ft_chi <- flextable(chi_summary) %>%
        theme_zebra() %>%
        autofit()
      
      doc <- doc %>% body_add_flextable(ft_chi)
      
      # Expected counts if requested
      if(input$chisq_expected && !is.null(results$test$expected)) {
        doc <- doc %>%
          body_add_par("EXPECTED COUNTS", style = "heading 2")
        
        expected_df <- as.data.frame.matrix(results$test$expected)
        expected_df <- cbind(Variable = rownames(expected_df), expected_df)
        
        ft_expected <- flextable(expected_df) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_expected)
      }
      
      # Residuals if requested
      if(input$chisq_residuals && !is.null(results$test$residuals)) {
        doc <- doc %>%
          body_add_par("PEARSON RESIDUALS", style = "heading 2")
        
        residuals_df <- as.data.frame.matrix(results$test$residuals)
        residuals_df <- cbind(Variable = rownames(residuals_df), residuals_df)
        
        ft_residuals <- flextable(residuals_df) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_residuals)
      }
      
      # Fisher's exact test if requested
      if(input$chisq_fisher && !is.null(results$fisher)) {
        doc <- doc %>%
          body_add_par("FISHER'S EXACT TEST", style = "heading 2")
        
        fisher_summary <- data.frame(
          Statistic = c("Odds Ratio", "P-value", "Confidence Interval", "Significance"),
          Value = c(
            ifelse(!is.null(results$fisher$estimate), round(results$fisher$estimate, 4), "N/A"),
            round(results$fisher$p.value, 4),
            ifelse(!is.null(results$fisher$conf.int), 
                   paste(round(results$fisher$conf.int[1], 4), "to", round(results$fisher$conf.int[2], 4)), 
                   "N/A"),
            ifelse(results$fisher$p.value < 0.05, "Statistically Significant", "Not Statistically Significant")
          )
        )
        
        ft_fisher <- flextable(fisher_summary) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_fisher)
      }
      
      # Relative risk if requested
      if(input$chisq_rr && !is.null(results$rr)) {
        doc <- doc %>%
          body_add_par("RELATIVE RISK ANALYSIS", style = "heading 2")
        
        # Extract relative risk information
        rr_data <- results$rr$data
        rr_estimate <- results$rr$measure
        rr_pvalue <- results$rr$p.value
        
        # Create RR summary table
        rr_summary <- data.frame(
          Statistic = c("Relative Risk", "P-value", "Significance"),
          Value = c(
            ifelse(!is.null(rr_estimate), round(rr_estimate[2, 1], 4), "N/A"),
            ifelse(!is.null(rr_pvalue), round(rr_pvalue[2, 1], 4), "N/A"),
            ifelse(!is.null(rr_pvalue) && rr_pvalue[2, 1] < 0.05, 
                   "Statistically Significant", "Not Statistically Significant")
          )
        )
        
        ft_rr <- flextable(rr_summary) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_rr)
      }
      
      # Interpretation section
      doc <- doc %>%
        body_add_par("INTERPRETATION", style = "heading 2")
      
      interpretation_text <- if (results$test$p.value < 0.05) {
        paste("The chi-square test reveals a statistically significant association between", 
              input$chisq_var1, "and", input$chisq_var2, 
              "(χ² =", round(results$test$statistic, 3), 
              ", df =", results$test$parameter, 
              ", p =", round(results$test$p.value, 4), ").",
              "This suggests that the two variables are not independent.")
      } else {
        paste("The chi-square test does not show a statistically significant association between", 
              input$chisq_var1, "and", input$chisq_var2, 
              "(χ² =", round(results$test$statistic, 3), 
              ", df =", results$test$parameter, 
              ", p =", round(results$test$p.value, 4), ").",
              "This suggests that any observed relationship could be due to chance.")
      }
      
      doc <- doc %>% 
        body_add_par(interpretation_text, style = "Normal") %>%
        body_add_par("", style = "Normal")
      
      # Data quality notes
      doc <- doc %>%
        body_add_par("DATA QUALITY NOTES", style = "heading 2")
      
      # Check for small expected counts
      if (!is.null(results$test$expected)) {
        small_expected <- sum(results$test$expected < 5)
        total_cells <- length(results$test$expected)
        
        quality_note <- if (small_expected > 0) {
          paste("Warning:", small_expected, "out of", total_cells, 
                "cells have expected counts less than 5. Consider using Fisher's exact test for more reliable results.")
        } else {
          "All expected cell counts are ≥5, meeting the assumptions for chi-square test."
        }
        
        doc <- doc %>% body_add_par(quality_note, style = "Normal")
      }
      
      print(doc, target = file)
    }
  )
  
  
  output$download_chisq_plot <- downloadHandler(
    filename = "chi_square_plot.png",
    content = function(file) {
      df <- processed_data()
      
      p <- ggplot(df, aes(x = !!sym(input$chisq_var1), fill = !!sym(input$chisq_var2))) +
        geom_bar(position = "fill") +
        labs(title = paste("Stacked Bar Plot:", input$chisq_var1, "vs", input$chisq_var2),
             x = input$chisq_var1,
             y = "Proportion",
             fill = input$chisq_var2) +
        theme_minimal() +
        theme(axis.text.x = element_text(angle = 45, hjust = 1))
      
      ggsave(file, plot = p, device = "png", width = 10, height = 6, dpi = 300)
    }
  )
  
  
  # Add to your data loading observer (around line 1070 in your original code)
  observeEvent(processed_data(), {
    req(processed_data())
    df_processed <- processed_data()
    
    # Update Poisson regression inputs
    updateSelectInput(session, "poisson_outcome", choices = names(df_processed)[sapply(df_processed, is.numeric)])
    updateSelectInput(session, "poisson_vars", choices = names(df_processed))
    updateSelectInput(session, "poisson_offset_var", choices = names(df_processed)[sapply(df_processed, is.numeric)])
    
    # Update central tendency inputs
    updateSelectInput(session, "central_var", choices = names(df_processed)[sapply(df_processed, is.numeric)])
    updateSelectInput(session, "central_stratify", choices = c("None" = "none", names(df_processed)))
    
    # Update ANOVA inputs - MOVED HERE to ensure df_processed exists
    updateSelectInput(session, "anova_outcome", choices = names(df_processed)[sapply(df_processed, is.numeric)])
    updateSelectInput(session, "anova_group", choices = names(df_processed))
    updateSelectInput(session, "anova_outcome_2way", choices = names(df_processed)[sapply(df_processed, is.numeric)])
    updateSelectInput(session, "anova_factor1", choices = names(df_processed))
    updateSelectInput(session, "anova_factor2", choices = names(df_processed))
    updateSelectInput(session, "welch_outcome", choices = names(df_processed)[sapply(df_processed, is.numeric)])
    updateSelectInput(session, "welch_group", choices = names(df_processed))
    updateSelectInput(session, "assumptions_outcome", choices = names(df_processed)[sapply(df_processed, is.numeric)])
    updateSelectInput(session, "assumptions_group", choices = names(df_processed))
    
    # Update T-test inputs
    updateSelectInput(session, "ttest_onesample_var", choices = names(df_processed)[sapply(df_processed, is.numeric)])
    updateSelectInput(session, "ttest_independent_var", choices = names(df_processed)[sapply(df_processed, is.numeric)])
    updateSelectInput(session, "ttest_independent_group", choices = names(df_processed))
    updateSelectInput(session, "ttest_paired_var1", choices = names(df_processed)[sapply(df_processed, is.numeric)])
    updateSelectInput(session, "ttest_paired_var2", choices = names(df_processed)[sapply(df_processed, is.numeric)])
    updateSelectInput(session, "mannwhitney_var", choices = names(df_processed)[sapply(df_processed, is.numeric)])
    updateSelectInput(session, "mannwhitney_group", choices = names(df_processed))
    updateSelectInput(session, "wilcoxon_var1", choices = names(df_processed)[sapply(df_processed, is.numeric)])
    updateSelectInput(session, "wilcoxon_var2", choices = names(df_processed)[sapply(df_processed, is.numeric)])
  })
  
  # UI for independent t-test group levels
  output$ttest_independent_levels_ui <- renderUI({
    req(input$ttest_independent_group, processed_data())
    df <- processed_data()
    
    if (!input$ttest_independent_group %in% names(df)) return(NULL)
    
    var_data <- df[[input$ttest_independent_group]]
    levels <- if (is.factor(var_data)) levels(var_data) else unique(na.omit(var_data))
    
    if (length(levels) != 2) {
      return(helpText("Grouping variable must have exactly 2 categories for independent t-test."))
    }
    
    tagList(
      selectInput("ttest_group1", "Group 1:", choices = levels, selected = levels[1]),
      selectInput("ttest_group2", "Group 2:", choices = levels, selected = levels[2])
    )
  })
  
  # One-Way ANOVA
  anova_results <- eventReactive(input$run_anova, {
    req(input$anova_outcome, input$anova_group, processed_data())
    
    tryCatch({
      df <- processed_data()
      
      # Validate variables
      if (!input$anova_outcome %in% names(df) || !input$anova_group %in% names(df)) {
        return(list(error = "Selected variables not found", success = FALSE))
      }
      
      outcome <- df[[input$anova_outcome]]
      group <- as.factor(df[[input$anova_group]])
      
      # Remove missing values
      complete_cases <- !is.na(outcome) & !is.na(group)
      outcome_clean <- outcome[complete_cases]
      group_clean <- group[complete_cases]
      
      if (length(unique(group_clean)) < 2) {
        return(list(error = "Grouping variable must have at least 2 levels", success = FALSE))
      }
      
      # Fit ANOVA model
      model <- aov(outcome_clean ~ group_clean)
      anova_summary <- summary(model)
      
      # Post-hoc tests if requested
      posthoc_results <- NULL
      if (input$anova_posthoc) {
        if (input$anova_posthoc_method == "tukey") {
          posthoc_results <- TukeyHSD(model)
        } else if (input$anova_posthoc_method == "bonferroni") {
          posthoc_results <- pairwise.t.test(outcome_clean, group_clean, p.adjust.method = "bonferroni")
        }
      }
      
      # Assumptions check if requested
      assumptions_results <- NULL
      if (input$anova_assumptions) {
        # Normality test
        normality <- shapiro.test(residuals(model))
        
        # Homogeneity of variance
        levene <- car::leveneTest(outcome_clean ~ group_clean)
        
        assumptions_results <- list(
          normality = normality,
          homogeneity = levene
        )
      }
      
      list(
        model = model,
        summary = anova_summary,
        posthoc = posthoc_results,
        assumptions = assumptions_results,
        success = TRUE
      )
      
    }, error = function(e) {
      return(list(error = paste("ANOVA error:", e$message), success = FALSE))
    })
  })
  
  output$anova_results <- renderPrint({
    req(anova_results())
    results <- anova_results()
    
    if (!results$success) {
      cat("ANOVA ERROR:\n")
      cat(results$error, "\n")
      return()
    }
    
    cat("ONE-WAY ANOVA RESULTS\n")
    cat("=====================\n\n")
    print(results$summary)
  })
  
  output$anova_assumptions_results <- renderPrint({
    req(anova_results())
    results <- anova_results()
    
    if (!results$success || is.null(results$assumptions)) return()
    
    cat("NORMALITY TEST (Shapiro-Wilk):\n")
    print(results$assumptions$normality)
    cat("\nHOMOGENEITY OF VARIANCE (Levene's Test):\n")
    print(results$assumptions$homogeneity)
  })
  
  output$anova_posthoc_results <- renderPrint({
    req(anova_results())
    results <- anova_results()
    
    if (!results$success || is.null(results$posthoc)) return()
    
    cat("POST-HOC COMPARISONS:\n")
    print(results$posthoc)
  })
  
  # Two-Way ANOVA
  anova_2way_results <- eventReactive(input$run_anova_2way, {
    req(input$anova_outcome_2way, input$anova_factor1, input$anova_factor2, processed_data())
    
    tryCatch({
      df <- processed_data()
      
      # Create formula
      if (input$anova_interaction) {
        formula <- as.formula(paste(input$anova_outcome_2way, "~", input$anova_factor1, "*", input$anova_factor2))
      } else {
        formula <- as.formula(paste(input$anova_outcome_2way, "~", input$anova_factor1, "+", input$anova_factor2))
      }
      
      model <- aov(formula, data = df)
      summary_result <- summary(model)
      
      # Simple effects analysis if requested
      simple_effects <- NULL
      if (input$anova_2way_posthoc && input$anova_interaction) {
        simple_effects <- emmeans(model, as.formula(paste("~", input$anova_factor1, "|", input$anova_factor2)))
      }
      
      list(
        model = model,
        summary = summary_result,
        simple_effects = simple_effects,
        success = TRUE
      )
      
    }, error = function(e) {
      return(list(error = paste("Two-way ANOVA error:", e$message), success = FALSE))
    })
  })
  
  output$anova_2way_results <- renderPrint({
    req(anova_2way_results())
    results <- anova_2way_results()
    
    if (!results$success) {
      cat("TWO-WAY ANOVA ERROR:\n")
      cat(results$error, "\n")
      return()
    }
    
    cat("TWO-WAY ANOVA RESULTS\n")
    cat("=====================\n\n")
    print(results$summary)
  })
  
  output$anova_2way_posthoc <- renderPrint({
    req(anova_2way_results())
    results <- anova_2way_results()
    
    if (!results$success || is.null(results$simple_effects)) return()
    
    cat("SIMPLE EFFECTS ANALYSIS:\n")
    print(results$simple_effects)
  })
  
  # Welch ANOVA
  welch_anova_results <- eventReactive(input$run_welch_anova, {
    req(input$welch_outcome, input$welch_group, processed_data())
    
    tryCatch({
      df <- processed_data()
      
      outcome <- df[[input$welch_outcome]]
      group <- as.factor(df[[input$welch_group]])
      
      # Remove missing values
      complete_cases <- !is.na(outcome) & !is.na(group)
      outcome_clean <- outcome[complete_cases]
      group_clean <- group[complete_cases]
      
      # Welch ANOVA
      welch_result <- oneway.test(outcome_clean ~ group_clean, var.equal = FALSE)
      
      # Games-Howell post-hoc
      posthoc <- tryCatch({
        PMCMRplus::gamesHowellTest(outcome_clean, group_clean)
      }, error = function(e) {
        NULL
      })
      
      list(
        welch_test = welch_result,
        posthoc = posthoc,
        success = TRUE
      )
      
    }, error = function(e) {
      return(list(error = paste("Welch ANOVA error:", e$message), success = FALSE))
    })
  })
  
  output$welch_anova_results <- renderPrint({
    req(welch_anova_results())
    results <- welch_anova_results()
    
    if (!results$success) {
      cat("WELCH ANOVA ERROR:\n")
      cat(results$error, "\n")
      return()
    }
    
    cat("WELCH ANOVA RESULTS\n")
    cat("===================\n\n")
    print(results$welch_test)
  })
  
  output$welch_posthoc_results <- renderPrint({
    req(welch_anova_results())
    results <- welch_anova_results()
    
    if (!results$success || is.null(results$posthoc)) {
      cat("Games-Howell post-hoc test not available.\n")
      return()
    }
    
    cat("GAMES-HOWELL POST-HOC TEST:\n")
    print(results$posthoc)
  })
  
  # One-Sample T-Test
  ttest_onesample_results <- eventReactive(input$run_ttest_onesample, {
    req(input$ttest_onesample_var, processed_data())
    
    tryCatch({
      df <- processed_data()
      var_data <- df[[input$ttest_onesample_var]]
      
      # Remove missing values
      var_clean <- var_data[!is.na(var_data)]
      
      if (length(var_clean) < 2) {
        return(list(error = "Insufficient data for t-test", success = FALSE))
      }
      
      # Perform t-test
      ttest <- t.test(var_clean, 
                      mu = input$ttest_onesample_mu,
                      alternative = input$ttest_onesample_alternative,
                      conf.level = input$ttest_onesample_conf)
      
      # Descriptive statistics
      descriptives <- data.frame(
        N = length(var_clean),
        Mean = mean(var_clean),
        SD = sd(var_clean),
        SE = sd(var_clean)/sqrt(length(var_clean)),
        Test_Value = input$ttest_onesample_mu
      )
      
      list(
        ttest = ttest,
        descriptives = descriptives,
        data = var_clean,
        success = TRUE
      )
      
    }, error = function(e) {
      return(list(error = paste("One-sample t-test error:", e$message), success = FALSE))
    })
  })
  
  output$ttest_onesample_results <- renderPrint({
    req(ttest_onesample_results())
    results <- ttest_onesample_results()
    
    if (!results$success) {
      cat("ONE-SAMPLE T-TEST ERROR:\n")
      cat(results$error, "\n")
      return()
    }
    
    cat("ONE-SAMPLE T-TEST RESULTS\n")
    cat("=========================\n\n")
    print(results$ttest)
  })
  
  output$ttest_onesample_descriptives <- renderPrint({
    req(ttest_onesample_results())
    results <- ttest_onesample_results()
    
    if (!results$success) return()
    
    cat("DESCRIPTIVE STATISTICS:\n")
    print(results$descriptives)
  })
  
  # Independent T-Test
  ttest_independent_results <- eventReactive(input$run_ttest_independent, {
    req(input$ttest_independent_var, input$ttest_independent_group, 
        input$ttest_group1, input$ttest_group2, processed_data())
    
    tryCatch({
      df <- processed_data()
      
      # Filter data for selected groups
      group_data <- df[[input$ttest_independent_group]]
      outcome_data <- df[[input$ttest_independent_var]]
      
      # Create subset for the two groups
      group1_data <- outcome_data[group_data == input$ttest_group1 & !is.na(outcome_data)]
      group2_data <- outcome_data[group_data == input$ttest_group2 & !is.na(outcome_data)]
      
      if (length(group1_data) < 2 || length(group2_data) < 2) {
        return(list(error = "Insufficient data in one or both groups", success = FALSE))
      }
      
      # Perform t-test
      ttest <- t.test(group1_data, group2_data,
                      var.equal = input$ttest_var_equal,
                      alternative = input$ttest_alternative,
                      conf.level = input$ttest_conf)
      
      # Variance check
      var_test <- var.test(group1_data, group2_data)
      
      # Group descriptives
      descriptives <- data.frame(
        Group = c(input$ttest_group1, input$ttest_group2),
        N = c(length(group1_data), length(group2_data)),
        Mean = c(mean(group1_data), mean(group2_data)),
        SD = c(sd(group1_data), sd(group2_data)),
        SE = c(sd(group1_data)/sqrt(length(group1_data)), 
               sd(group2_data)/sqrt(length(group2_data)))
      )
      
      list(
        ttest = ttest,
        var_test = var_test,
        descriptives = descriptives,
        group1_data = group1_data,
        group2_data = group2_data,
        success = TRUE
      )
      
    }, error = function(e) {
      return(list(error = paste("Independent t-test error:", e$message), success = FALSE))
    })
  })
  
  output$ttest_independent_results <- renderPrint({
    req(ttest_independent_results())
    results <- ttest_independent_results()
    
    if (!results$success) {
      cat("INDEPENDENT T-TEST ERROR:\n")
      cat(results$error, "\n")
      return()
    }
    
    cat("INDEPENDENT T-TEST RESULTS\n")
    cat("==========================\n\n")
    print(results$ttest)
  })
  
  output$ttest_independent_descriptives <- renderPrint({
    req(ttest_independent_results())
    results <- ttest_independent_results()
    
    if (!results$success) return()
    
    cat("GROUP DESCRIPTIVES:\n")
    print(results$descriptives)
  })
  
  output$ttest_variance_check <- renderPrint({
    req(ttest_independent_results())
    results <- ttest_independent_results()
    
    if (!results$success) return()
    
    cat("VARIANCE EQUALITY TEST (F-test):\n")
    print(results$var_test)
    cat("\nInterpretation: ", ifelse(results$var_test$p.value < 0.05, 
                                     "Variances are significantly different (use Welch test)",
                                     "Variances are not significantly different"))
  })
  
  # Paired T-Test
  ttest_paired_results <- eventReactive(input$run_ttest_paired, {
    req(input$ttest_paired_var1, input$ttest_paired_var2, processed_data())
    
    tryCatch({
      df <- processed_data()
      
      var1_data <- df[[input$ttest_paired_var1]]
      var2_data <- df[[input$ttest_paired_var2]]
      
      # Remove pairs with missing values
      complete_pairs <- !is.na(var1_data) & !is.na(var2_data)
      var1_clean <- var1_data[complete_pairs]
      var2_clean <- var2_data[complete_pairs]
      
      if (length(var1_clean) < 2) {
        return(list(error = "Insufficient paired data", success = FALSE))
      }
      
      # Perform paired t-test
      ttest <- t.test(var1_clean, var2_clean,
                      paired = TRUE,
                      alternative = input$ttest_paired_alternative,
                      conf.level = input$ttest_paired_conf)
      
      # Difference analysis
      differences <- var1_clean - var2_clean
      
      descriptives <- data.frame(
        Variable = c(input$ttest_paired_var1, input$ttest_paired_var2, "Difference"),
        N = c(length(var1_clean), length(var2_clean), length(differences)),
        Mean = c(mean(var1_clean), mean(var2_clean), mean(differences)),
        SD = c(sd(var1_clean), sd(var2_clean), sd(differences))
      )
      
      list(
        ttest = ttest,
        descriptives = descriptives,
        differences = differences,
        var1_data = var1_clean,
        var2_data = var2_clean,
        success = TRUE
      )
      
    }, error = function(e) {
      return(list(error = paste("Paired t-test error:", e$message), success = FALSE))
    })
  })
  
  output$ttest_paired_results <- renderPrint({
    req(ttest_paired_results())
    results <- ttest_paired_results()
    
    if (!results$success) {
      cat("PAIRED T-TEST ERROR:\n")
      cat(results$error, "\n")
      return()
    }
    
    cat("PAIRED T-TEST RESULTS\n")
    cat("=====================\n\n")
    print(results$ttest)
  })
  
  output$ttest_paired_descriptives <- renderPrint({
    req(ttest_paired_results())
    results <- ttest_paired_results()
    
    if (!results$success) return()
    
    cat("PAIRED DESCRIPTIVES:\n")
    print(results$descriptives)
  })
  
  output$ttest_paired_differences <- renderPrint({
    req(ttest_paired_results())
    results <- ttest_paired_results()
    
    if (!results$success) return()
    
    cat("DIFFERENCE ANALYSIS:\n")
    cat("Correlation between pairs:", cor(results$var1_data, results$var2_data), "\n")
    cat("Mean difference:", mean(results$differences), "\n")
    cat("SD of differences:", sd(results$differences), "\n")
  })
  
  # Non-parametric Tests
  mannwhitney_results <- eventReactive(input$run_mannwhitney, {
    req(input$mannwhitney_var, input$mannwhitney_group, processed_data())
    
    tryCatch({
      df <- processed_data()
      
      outcome <- df[[input$mannwhitney_var]]
      group <- as.factor(df[[input$mannwhitney_group]])
      
      # Remove missing values
      complete_cases <- !is.na(outcome) & !is.na(group)
      outcome_clean <- outcome[complete_cases]
      group_clean <- group[complete_cases]
      
      if (length(unique(group_clean)) != 2) {
        return(list(error = "Mann-Whitney test requires exactly 2 groups", success = FALSE))
      }
      
      test_result <- wilcox.test(outcome_clean ~ group_clean, exact = FALSE)
      
      list(
        test = test_result,
        success = TRUE
      )
      
    }, error = function(e) {
      return(list(error = paste("Mann-Whitney test error:", e$message), success = FALSE))
    })
  })
  
  output$mannwhitney_results <- renderPrint({
    req(mannwhitney_results())
    results <- mannwhitney_results()
    
    if (!results$success) {
      cat("MANN-WHITNEY TEST ERROR:\n")
      cat(results$error, "\n")
      return()
    }
    
    cat("MANN-WHITNEY U TEST RESULTS\n")
    cat("===========================\n\n")
    print(results$test)
  })
  
  wilcoxon_results <- eventReactive(input$run_wilcoxon, {
    req(input$wilcoxon_var1, input$wilcoxon_var2, processed_data())
    
    tryCatch({
      df <- processed_data()
      
      var1 <- df[[input$wilcoxon_var1]]
      var2 <- df[[input$wilcoxon_var2]]
      
      # Remove pairs with missing values
      complete_pairs <- !is.na(var1) & !is.na(var2)
      var1_clean <- var1[complete_pairs]
      var2_clean <- var2[complete_pairs]
      
      if (length(var1_clean) < 2) {
        return(list(error = "Insufficient paired data", success = FALSE))
      }
      
      test_result <- wilcox.test(var1_clean, var2_clean, paired = TRUE, exact = FALSE)
      
      list(
        test = test_result,
        success = TRUE
      )
      
    }, error = function(e) {
      return(list(error = paste("Wilcoxon test error:", e$message), success = FALSE))
    })
  })
  
  output$wilcoxon_results <- renderPrint({
    req(wilcoxon_results())
    results <- wilcoxon_results()
    
    if (!results$success) {
      cat("WILCOXON TEST ERROR:\n")
      cat(results$error, "\n")
      return()
    }
    
    cat("WILCOXON SIGNED-RANK TEST RESULTS\n")
    cat("=================================\n\n")
    print(results$test)
  })
  
  
  # Navigation buttons
  observeEvent(input$goto_data, {
    updateNavbarPage(session, "nav", selected = "Data Upload")
  })
  
  observeEvent(input$goto_descriptive, {
    updateNavbarPage(session, "nav", selected = "Descriptive Analysis")
  })
  
  observeEvent(input$goto_disease, {
    updateNavbarPage(session, "nav", selected = "Disease Frequency")
  })
  
  observeEvent(input$goto_statistics, {
    updateNavbarPage(session, "nav", selected = "Statistical Associations")
  })
}


# Run the application
shinyApp(ui = ui, server = server)
