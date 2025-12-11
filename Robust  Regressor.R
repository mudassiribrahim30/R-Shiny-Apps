# app.R

library(shiny)
library(MASS)
library(robustbase)
library(DT)
library(tidyverse)
library(officer)
library(lm.beta)
library(readxl)
library(haven)

ui <- fluidPage(
  tags$head(
    tags$style(HTML("
      .large-output pre {
        font-size: 18px;
      }
      .shiny-output-error-validation {
        color: red;
        font-weight: bold;
      }
      table.dataTable tbody td {
        font-size: 16px;
      }
      table.dataTable thead th {
        font-size: 17px;
        font-weight: bold;
      }
      .well-panel {
        background-color: #f5f5f5;
        border-radius: 5px;
        padding: 15px;
        margin-bottom: 15px;
      }
      .ref-cat-box {
        background-color: #f8f9fa;
        border-left: 4px solid #6c757d;
        padding: 10px;
        margin-bottom: 10px;
      }
      .logo-container {
        text-align: center;
        margin-bottom: 20px;
      }
      .logo-img {
        max-width: 200px;
        max-height: 100px;
      }
      .about-section {
        padding: 20px;
        background-color: #f9f9f9;
        border-radius: 5px;
        margin-top: 20px;
      }
      .references {
        font-size: 14px;
        margin-top: 20px;
      }
      .interpretation {
        background-color: #f0f8ff;
        padding: 15px;
        border-radius: 5px;
        margin-top: 15px;
        border-left: 4px solid #4682b4;
      }
      .signif-yes {
        background-color: #e6ffe6 !important;
      }
      .signif-no {
        background-color: #ffffff !important;
      }
    "))
  ),
  
  div(class = "logo-container",
      img(src = "https://cdn-icons-png.flaticon.com/512/2103/2103633.png", 
          class = "logo-img", alt = "Robust Regressor Logo")
  ),
  
  titlePanel(
    div(
      h1("Robust Regressor", style = "color: #2c3e50; text-align: center;"),
      p("A tool for robust regression analysis with categorical variable support", 
        style = "color: #7f8c8d; font-style: italic; text-align: center;")
    )
  ),
  
  sidebarLayout(
    sidebarPanel(
      div(class = "well-panel",
          h4("Data Input"),
          fileInput("datafile", "Upload File (CSV, Excel, Stata, SPSS)", 
                    accept = c(".csv", ".xlsx", ".xls", ".dta", ".sav", ".zsav", ".por"),
                    buttonLabel = "Browse...",
                    placeholder = "No file selected")
      ),
      
      uiOutput("varselect_ui"),
      
      uiOutput("ref_cat_ui"),
      
      div(class = "well-panel",
          h4("Model Settings"),
          selectInput("package", "Select Package", 
                      choices = c("robustbase", "MASS"),
                      width = "100%"),
          
          conditionalPanel(
            condition = "input.package == 'robustbase'",
            selectInput("method_robustbase", "Select Method (lmrob)",
                        choices = c("MM", "M", "S"), selected = "MM",
                        width = "100%")
          ),
          
          conditionalPanel(
            condition = "input.package == 'MASS'",
            selectInput("method_mass", "Select Method (rlm)",
                        choices = c("M", "MM"), selected = "M",
                        width = "100%")
          ),
          
          numericInput("decimal_places", "Decimal Places", 
                       value = 4, min = 1, max = 6, width = "100%")
      ),
      
      actionButton("run_analysis", "Run Analysis", 
                   class = "btn btn-primary",
                   style = "width: 100%; font-weight: bold;"),
      br(),
      textOutput("na_message")
    ),
    
    mainPanel(
      tabsetPanel(
        tabPanel("Model Summary", 
                 icon = icon("chart-line"),
                 div(class = "large-output", verbatimTextOutput("model_summary")),
                 uiOutput("model_interpretation")),
        tabPanel("Estimates Table", 
                 icon = icon("table"),
                 DTOutput("std_table"),
                 uiOutput("table_interpretation")),
        tabPanel("Data View", 
                 icon = icon("database"),
                 h4("Data Preview"),
                 DTOutput("datatable")),
        tabPanel("Download Full Results", 
                 icon = icon("file-download"),
                 downloadButton("download_full", "Download Full Results (Word)",
                                class = "btn btn-success")),
        tabPanel("About",
                 icon = icon("info-circle"),
                 div(class = "about-section",
                     h3("About Robust Regressor"),
                     p("Robust Regressor is a user-friendly Shiny application designed for performing robust regression analysis. It leverages two widely used R packages (MASS and robustbase) to provide reliable estimates that are resistant to outliers and data irregularities."),
                     
                     h4("When to Use Robust Regression vs Ordinary Least Squares (OLS)"),
                     p("Robust regression should be used when:"),
                     tags$ul(
                       tags$li("Your data contains outliers that may influence the results"),
                       tags$li("The assumption of normally distributed errors is violated"),
                       tags$li("You suspect heteroscedasticity (non-constant variance)"),
                       tags$li("You want to reduce the influence of leverage points")
                     ),
                     p("Compared to OLS, robust regression:"),
                     tags$ul(
                       tags$li("Provides more reliable estimates when outliers are present"),
                       tags$li("Is less sensitive to violations of normality assumptions"),
                       tags$li("Gives more weight to typical observations than outliers"),
                       tags$li("May have slightly less statistical efficiency with clean data")
                     ),
                     
                     h4("Interpreting Multivariate Results"),
                     p("In multivariate robust regression:"),
                     tags$ul(
                       tags$li("Each variable's significance is evaluated independently"),
                       tags$li("A variable may be significant even when others are not"),
                       tags$li("The robust procedure reduces but doesn't eliminate multicollinearity effects"),
                       tags$li("Examine both unstandardized and standardized coefficients for complete understanding")
                     ),
                     
                     div(class = "references",
                         h4("References"),
                         p("Maechler, M., Rousseeuw, P., Croux, C., Todorov, V., Ruckstuhl, A., Salibian-Barrera, M., ... & Verbeke, T. (2021). robustbase: Basic Robust Statistics. R package version 0.93-9."),
                         p("Venables, W. N. & Ripley, B. D. (2002) Modern Applied Statistics with S. Fourth Edition. Springer, New York. ISBN 0-387-95457-0"),
                         p("Huber, P. J. (1981). Robust Statistics. Wiley."),
                         
                         h4("Developer Information"),
                         p("Developed by Mudasir Mohammed Ibrahim"),
                         p("For suggestions or problems, please contact: mudassiribrahim30@gmail.com")
                     )
                 )
        )
      )
    )
  )
)

server <- function(input, output, session) {
  
  # Function to safely read data with proper encoding
  safe_read_data <- function(path, ext) {
    tryCatch({
      df <- switch(ext,
                   csv = read_csv(path, show_col_types = FALSE, 
                                  locale = locale(encoding = "UTF-8")),
                   xlsx = read_excel(path),
                   xls = read_excel(path),
                   dta = read_dta(path),
                   sav = read_sav(path),
                   zsav = read_sav(path),
                   por = read_por(path),
                   stop("Unsupported file type")
      )
      
      # Clean column names - replace spaces with underscores and remove special characters
      names(df) <- gsub("[^[:alnum:]_\\.]", "_", names(df))
      names(df) <- gsub("_{2,}", "_", names(df))  # Remove multiple underscores
      names(df) <- gsub("^_|_$", "", names(df))   # Remove leading/trailing underscores
      
      return(df)
    }, error = function(e) {
      showNotification(paste("Error reading file:", e$message), type = "error")
      return(NULL)
    })
  }
  
  dataset <- reactive({
    req(input$datafile)
    ext <- tools::file_ext(input$datafile$name)
    
    df <- safe_read_data(input$datafile$datapath, tolower(ext))
    
    if (is.null(df)) {
      validate("Failed to read the file. Please check file format and encoding.")
    }
    
    validate(
      need(ncol(df) > 0, "File appears to be empty or has no columns."),
      need(nrow(df) > 0, "File appears to be empty or has no rows."),
      need(ncol(df) <= 500, "File must contain 500 or fewer variables."),
      need(nrow(df) <= 10000, "File must contain 10,000 or fewer observations.")
    )
    
    # Convert character variables to factors
    df <- df %>% 
      mutate(across(where(is.character), as.factor)) %>%
      # Remove completely empty rows
      filter(if_any(everything(), ~!is.na(.)))
    
    # Store original names for display
    attr(df, "original_names") <- names(df)
    
    df
  })
  
  # Store original variable names for display
  original_names <- reactive({
    req(dataset())
    attr(dataset(), "original_names")
  })
  
  output$varselect_ui <- renderUI({
    req(dataset())
    vars <- names(dataset())
    orig_names <- original_names()
    
    # Create display names with both original and cleaned names
    display_names <- ifelse(vars == orig_names, 
                            vars, 
                            paste0(orig_names, " [", vars, "]"))
    
    tagList(
      div(class = "well-panel",
          h4("Variable Selection"),
          selectInput("depvar", "Dependent Variable", 
                      choices = setNames(vars, display_names), 
                      width = "100%"),
          selectInput("indepvars", "Independent Variable(s)", 
                      choices = setNames(vars, display_names), 
                      multiple = TRUE, width = "100%")
      )
    )
  })
  
  # Function to safely create formula with backticks for non-standard names
  create_safe_formula <- function(depvar, indepvars) {
    # Add backticks around variable names if they contain special characters
    safe_depvar <- ifelse(grepl("[^[:alnum:]_]", depvar), 
                          paste0("`", depvar, "`"), 
                          depvar)
    
    safe_indepvars <- sapply(indepvars, function(var) {
      ifelse(grepl("[^[:alnum:]_]", var), 
             paste0("`", var, "`"), 
             var)
    })
    
    formula_str <- paste(safe_depvar, "~", paste(safe_indepvars, collapse = " + "))
    as.formula(formula_str)
  }
  
  # Identify categorical variables from selected independent variables
  cat_vars <- reactive({
    req(input$indepvars, dataset())
    indepvars <- input$indepvars
    data <- dataset()
    
    # Get categorical variables (factors or character)
    cat_vars <- names(data)[sapply(data, function(x) is.factor(x) || is.character(x))]
    
    # Only return those that are in selected independent variables
    intersect(indepvars, cat_vars)
  })
  
  # UI for reference category selection
  output$ref_cat_ui <- renderUI({
    req(cat_vars(), dataset())
    vars <- cat_vars()
    
    if (length(vars) > 0) {
      ref_boxes <- lapply(vars, function(var) {
        var_data <- dataset()[[var]]
        choices <- if (is.factor(var_data)) {
          levels(var_data)
        } else {
          unique(na.omit(var_data))
        }
        
        # Get display name
        orig_names <- original_names()
        var_index <- which(names(dataset()) == var)
        display_name <- if (length(var_index) > 0) orig_names[var_index] else var
        
        div(class = "ref-cat-box",
            selectInput(paste0("ref_", var), 
                        label = paste("Reference Category for:", display_name), 
                        choices = choices,
                        selected = choices[1],
                        width = "100%")
        )
      })
      
      div(class = "well-panel",
          h4("Reference Categories"),
          ref_boxes
      )
    }
  })
  
  # Reactive value to store processed data with correct reference categories
  processed_data <- reactive({
    req(input$depvar, input$indepvars, dataset())
    
    selected_vars <- c(input$depvar, input$indepvars)
    data <- dataset() %>% select(all_of(selected_vars))
    
    # Set reference levels for categorical variables
    vars <- cat_vars()
    for (var in vars) {
      ref_level <- input[[paste0("ref_", var)]]
      if (!is.null(ref_level)) {
        var_data <- data[[var]]
        
        # Convert to factor if not already
        if (!is.factor(var_data)) {
          var_data <- as.factor(var_data)
        }
        
        current_levels <- levels(var_data)
        if (is.null(current_levels)) {
          current_levels <- unique(na.omit(var_data))
        }
        
        # Ensure ref_level exists in the data
        if (ref_level %in% current_levels) {
          new_levels <- c(ref_level, setdiff(current_levels, ref_level))
          data[[var]] <- factor(var_data, levels = new_levels)
        }
      }
    }
    
    data
  })
  
  model_data <- reactiveVal(NULL)
  
  observeEvent(input$run_analysis, {
    req(processed_data())
    
    data <- processed_data()
    
    # Check if data has enough observations
    validate(
      need(nrow(data) > 0, "No data available after processing."),
      need(nrow(data) >= length(input$indepvars) + 1, 
           paste("Not enough observations for the number of predictors.",
                 "Need at least", length(input$indepvars) + 1, "observations."))
    )
    
    na_rows <- sum(!complete.cases(data))
    clean_data <- na.omit(data)
    
    # Check if enough observations remain
    validate(
      need(nrow(clean_data) > 0, 
           "No complete cases available after removing missing values."),
      need(nrow(clean_data) >= length(input$indepvars) + 1,
           paste("Not enough complete cases for analysis.",
                 "Need at least", length(input$indepvars) + 1, "complete observations."))
    )
    
    model_data(clean_data)
    
    if (na_rows > 0) {
      output$na_message <- renderText({
        paste(na_rows, "rows with missing values were excluded from the analysis.")
      })
    } else {
      output$na_message <- renderText({ "No missing values detected." })
    }
  })
  
  model_result <- reactive({
    req(model_data())
    
    data <- model_data()
    depvar <- input$depvar
    indepvars <- input$indepvars
    
    # Safely create formula with backticks for non-standard names
    formula <- create_safe_formula(depvar, indepvars)
    
    if (input$package == "robustbase") {
      model <- tryCatch({
        lmrob(formula, data = data, method = input$method_robustbase)
      }, error = function(e) {
        showNotification(paste("Error in robustbase model:", e$message), 
                         type = "error", duration = 10)
        return(NULL)
      })
    } else {
      model <- tryCatch({
        rlm(formula, data = data, method = input$method_mass)
      }, error = function(e) {
        showNotification(paste("Error in MASS model:", e$message), 
                         type = "error", duration = 10)
        return(NULL)
      })
    }
    
    validate(need(!is.null(model), "Model fitting failed. Check your data and settings."))
    model
  })
  
  # Function to get standardized coefficients
  get_standardized_coefs <- function(model) {
    if (inherits(model, "lmrob")) {
      # For robustbase models, create temporary lm object
      temp_lm <- tryCatch({
        lm(model$terms, data = model$model)
      }, error = function(e) {
        return(NULL)
      })
      
      if (!is.null(temp_lm)) {
        lm.beta(temp_lm)$standardized.coefficients
      } else {
        rep(NA, length(coef(model)))
      }
    } else {
      # For MASS rlm models
      beta_result <- tryCatch({
        lm.beta(model)$standardized.coefficients
      }, error = function(e) {
        return(NULL)
      })
      
      if (!is.null(beta_result)) {
        beta_result
      } else {
        rep(NA, length(coef(model)))
      }
    }
  }
  
  # Function to extract coefficients and p-values from model summary
  get_model_coefficients <- function(model) {
    tryCatch({
      summ <- summary(model)
      coef_table <- summ$coefficients
      
      if (inherits(model, "lmrob")) {
        # For robustbase models, use the p-values directly from summary
        data.frame(
          Term = rownames(coef_table),
          Estimate = coef_table[, 1],
          Std.Error = coef_table[, 2],
          Statistic = coef_table[, 3],
          p_value = coef_table[, 4],
          stringsAsFactors = FALSE
        )
      } else {
        # For MASS rlm models, calculate approximate p-values from t-values
        df <- model$df.residual
        p_value <- 2 * pt(abs(coef_table[, 3]), df = df, lower.tail = FALSE)
        data.frame(
          Term = rownames(coef_table),
          Estimate = coef_table[, 1],
          Std.Error = coef_table[, 2],
          Statistic = coef_table[, 3],
          p_value = p_value,
          stringsAsFactors = FALSE
        )
      }
    }, error = function(e) {
      # Return minimal data frame if summary fails
      coefs <- coef(model)
      data.frame(
        Term = names(coefs),
        Estimate = coefs,
        Std.Error = rep(NA, length(coefs)),
        Statistic = rep(NA, length(coefs)),
        p_value = rep(NA, length(coefs)),
        stringsAsFactors = FALSE
      )
    })
  }
  
  output$model_summary <- renderPrint({
    req(model_result())
    print(summary(model_result()))
  })
  
  # Interpretation of model summary
  output$model_interpretation <- renderUI({
    req(model_result())
    
    model <- model_result()
    is_robustbase <- input$package == "robustbase"
    
    interpretation_content <- if (is_robustbase) {
      tagList(
        p("For robustbase (lmrob) models:"),
        tags$ul(
          tags$li("The 'Pr(>|t|)' column shows exact p-values from the robust model"),
          tags$li("Variables with p < 0.05 are considered statistically significant"),
          tags$li("Each variable's significance is evaluated independently in the multivariate context"),
          tags$li("The estimates are robust to outliers and influential points")
        )
      )
    } else {
      tagList(
        p("For MASS (rlm) models:"),
        tags$ul(
          tags$li("No exact p-values are provided - using t-value approximations"),
          tags$li("Variables with |t| > 2 are considered likely significant"),
          tags$li("Each variable's contribution is evaluated independently"),
          tags$li("The estimates are resistant to outliers in the response")
        )
      )
    }
    
    div(class = "interpretation",
        h4("Interpretation Guide"),
        interpretation_content,
        p("For both models:"),
        tags$ul(
          tags$li("Examine the confidence intervals for precision of estimates"),
          tags$li("Compare standardized coefficients to assess relative importance"),
          tags$li("Check model convergence/weights for potential issues")
        )
    )
  })
  
  output$std_table <- renderDT({
    req(model_result())
    
    model <- model_result()
    coef_df <- get_model_coefficients(model)
    
    # Get standardized coefficients
    std_coefs <- get_standardized_coefs(model)
    
    # Create data frame with all needed information
    df <- coef_df %>%
      mutate(
        Std_Estimate = std_coefs[match(Term, names(std_coefs))],
        Lower_95_CI = Estimate - 1.96 * Std.Error,
        Upper_95_CI = Estimate + 1.96 * Std.Error,
        Significance = if (input$package == "robustbase") {
          ifelse(p_value < 0.05 & !is.na(p_value), 
                 paste0("Significant (p = ", format.pval(p_value, digits = 3), ")"), 
                 ifelse(is.na(p_value), "NA", 
                        paste0("Not Significant (p = ", format.pval(p_value, digits = 3), ")")))
        } else {
          ifelse(abs(Statistic) > 2 & !is.na(Statistic), 
                 paste0("Likely Significant (|t| = ", round(abs(Statistic), 2), ")"), 
                 ifelse(is.na(Statistic), "NA", 
                        paste0("Not Significant (|t| = ", round(abs(Statistic), 2), ")")))
        },
        across(where(is.numeric), ~round(.x, input$decimal_places))
      ) %>%
      select(Term, Estimate, Std_Estimate, Std.Error, Lower_95_CI, Upper_95_CI, Significance)
    
    # Create the datatable
    dt <- DT::datatable(
      df,
      options = list(
        dom = 't',
        pageLength = nrow(df),
        scrollX = TRUE
      ),
      rownames = FALSE,
      caption = "Model Estimates with Standardized Coefficients and 95% Confidence Intervals",
      colnames = c('Term', 'Unstandardized Estimate', 'Standardized Estimate', 'Std. Error', 
                   'Lower 95% CI', 'Upper 95% CI', 'Significance')
    ) %>%
      formatStyle(names(df), fontSize = '16px')
    
    # Apply conditional formatting based on significance
    if (input$package == "robustbase") {
      sig_values <- unique(df$Significance)[str_detect(unique(df$Significance), "Significant")]
      dt <- dt %>%
        formatStyle('Significance',
                    backgroundColor = styleEqual(
                      sig_values,
                      rep('#e6ffe6', length(sig_values))
                    ))
    } else {
      sig_values <- unique(df$Significance)[str_detect(unique(df$Significance), "Likely Significant")]
      dt <- dt %>%
        formatStyle('Significance',
                    backgroundColor = styleEqual(
                      sig_values,
                      rep('#e6ffe6', length(sig_values))
                    ))
    }
    
    dt
  })
  
  # Interpretation of estimates table
  output$table_interpretation <- renderUI({
    req(model_result())
    
    is_robustbase <- input$package == "robustbase"
    
    interpretation_content <- if (is_robustbase) {
      tags$ul(
        tags$li(tags$b("Significant:"), " Variables with p < 0.05 (green highlight) make a statistically significant contribution"),
        tags$li(tags$b("Exact p-values:"), " Shown in parentheses for each variable"),
        tags$li(tags$b("Unstandardized Estimate:"), " The change in the dependent variable for a one-unit change in the predictor"),
        tags$li(tags$b("Standardized Estimate:"), " Allows comparison of effect sizes across variables")
      )
    } else {
      tags$ul(
        tags$li(tags$b("Likely Significant:"), " Variables with |t| > 2 (green highlight) likely make a significant contribution"),
        tags$li(tags$b("t-values:"), " Shown in parentheses for each variable"),
        tags$li(tags$b("Unstandardized Estimate:"), " The change in the dependent variable for a one-unit change in the predictor"),
        tags$li(tags$b("Standardized Estimate:"), " Allows comparison of effect sizes across variables")
      )
    }
    
    div(class = "interpretation",
        h4("Interpreting the Estimates Table"),
        interpretation_content,
        p("For both methods, examine confidence intervals to assess precision of estimates.")
    )
  })
  
  output$datatable <- renderDT({
    req(dataset())
    DT::datatable(
      head(dataset(), 50), 
      options = list(scrollX = TRUE, pageLength = 10),
      class = 'cell-border stripe',
      colnames = original_names()  # Show original names in display
    )
  })
  
  output$download_full <- downloadHandler(
    filename = function() { 
      paste("Robust_Regression_Full_Results_", Sys.Date(), ".docx", sep = "") 
    },
    content = function(file) {
      req(model_result())
      
      model <- model_result()
      
      # Get the main results
      coef_df <- get_model_coefficients(model)
      std_coefs <- get_standardized_coefs(model)
      
      # Create formatted results table
      results_df <- coef_df %>%
        mutate(
          Std_Estimate = std_coefs[match(Term, names(std_coefs))],
          Lower_CI = Estimate - 1.96 * Std.Error,
          Upper_CI = Estimate + 1.96 * Std.Error,
          CI = paste0("[", round(Lower_CI, input$decimal_places), ", ", 
                      round(Upper_CI, input$decimal_places), "]"),
          p_value_formatted = if (input$package == "robustbase") {
            ifelse(is.na(p_value), "NA", format.pval(p_value, digits = 3, eps = 0.001))
          } else {
            ifelse(is.na(Statistic), "NA", paste0("t = ", round(Statistic, 2)))
          }
        ) %>%
        select(Term, Estimate, Std.Error, Std_Estimate, CI, p_value_formatted)
      
      # Create a new Word document
      doc <- read_docx() %>%
        body_add_par("ROBUST REGRESSION ANALYSIS REPORT", style = "heading 1") %>%
        body_add_par("") %>%
        
        # Add key information section
        body_add_par("KEY INFORMATION", style = "heading 2") %>%
        body_add_par("") %>%
        body_add_par(paste("Date of Analysis:", Sys.Date()), style = "Normal") %>%
        body_add_par(paste("Dependent Variable:", input$depvar), style = "Normal") %>%
        body_add_par(paste("Independent Variables:", paste(input$indepvars, collapse = ", ")), style = "Normal") %>%
        body_add_par(paste("Package Used:", input$package), style = "Normal") %>%
        body_add_par(paste("Method:", 
                           ifelse(input$package == "robustbase", 
                                  input$method_robustbase, 
                                  input$method_mass)), style = "Normal") %>%
        body_add_par(paste("Sample Size (after removing missing):", nrow(model_data())), style = "Normal") %>%
        body_add_par("") %>%
        
        # Add categorical variable reference levels
        body_add_par("REFERENCE CATEGORIES FOR CATEGORICAL VARIABLES", style = "heading 2") %>%
        body_add_par("")
      
      # Add reference category information
      cat_vars_list <- cat_vars()
      if (length(cat_vars_list) > 0) {
        for (var in cat_vars_list) {
          ref_level <- input[[paste0("ref_", var)]]
          if (!is.null(ref_level)) {
            doc <- doc %>% 
              body_add_par(paste(var, "-> Reference Category:", ref_level), style = "Normal")
          }
        }
      } else {
        doc <- doc %>% 
          body_add_par("No categorical variables selected or all variables are continuous.", style = "Normal")
      }
      
      doc <- doc %>%
        body_add_par("") %>%
        
        # Add model summary
        body_add_par("MODEL SUMMARY", style = "heading 2") %>%
        body_add_par("")
      
      # Capture model summary as text
      model_summary_text <- capture.output(summary(model))
      for (line in model_summary_text) {
        doc <- doc %>% body_add_par(line, style = "Normal")
      }
      
      doc <- doc %>%
        body_add_par("") %>%
        
        # Add detailed results table
        body_add_par("DETAILED COEFFICIENTS TABLE", style = "heading 2") %>%
        body_add_par("")
      
      # Create a simple table using body_add_table (not flextable)
      # First, format the table nicely
      formatted_table <- results_df %>%
        mutate(
          Estimate = round(Estimate, input$decimal_places),
          Std.Error = round(Std.Error, input$decimal_places),
          Std_Estimate = round(Std_Estimate, input$decimal_places)
        ) %>%
        rename(
          Predictor = Term,
          b = Estimate,
          SE = Std.Error,
          β = Std_Estimate,
          `95% CI` = CI,
          `Significance Test` = p_value_formatted
        )
      
      # Add the table to the document
      doc <- doc %>% 
        body_add_table(
          formatted_table,
          header = TRUE,
          style = "table_template"
        ) %>%
        body_add_par("") %>%
        body_add_par("Note: b = unstandardized coefficient, β = standardized coefficient, SE = standard error, CI = confidence interval.", style = "Normal") %>%
        body_add_par("") %>%
        
        # Add interpretation section
        body_add_par("INTERPRETATION GUIDELINES", style = "heading 2") %>%
        body_add_par("")
      
      # Add interpretation based on package used
      if (input$package == "robustbase") {
        doc <- doc %>%
          body_add_par("For robustbase (lmrob) models:", style = "Normal") %>%
          body_add_par("• Variables with p < 0.05 are statistically significant", style = "Normal") %>%
          body_add_par("• The estimates are robust to outliers and influential points", style = "Normal") %>%
          body_add_par("• Examine both unstandardized (b) and standardized (β) coefficients", style = "Normal") %>%
          body_add_par("")
      } else {
        doc <- doc %>%
          body_add_par("For MASS (rlm) models:", style = "Normal") %>%
          body_add_par("• Variables with |t| > 2 are likely significant", style = "Normal") %>%
          body_add_par("• The estimates are resistant to outliers in the response", style = "Normal") %>%
          body_add_par("• Examine both unstandardized (b) and standardized (β) coefficients", style = "Normal") %>%
          body_add_par("")
      }
      
      # Add general interpretation
      doc <- doc %>%
        body_add_par("General Interpretation:", style = "Normal") %>%
        body_add_par("1. Unstandardized coefficients (b) show the change in the dependent variable for a one-unit change in the predictor", style = "Normal") %>%
        body_add_par("2. Standardized coefficients (β) allow comparison of effect sizes across different variables", style = "Normal") %>%
        body_add_par("3. Confidence intervals show the precision of the estimates", style = "Normal") %>%
        body_add_par("4. Smaller standard errors indicate more precise estimates", style = "Normal") %>%
        body_add_par("") %>%
        
        # Add model diagnostics
        body_add_par("MODEL DIAGNOSTICS", style = "heading 2") %>%
        body_add_par("") %>%
        body_add_par(paste("Residual degrees of freedom:", model$df.residual), style = "Normal") %>%
        body_add_par(paste("Number of observations:", nrow(model_data())), style = "Normal")
      
      # Check model convergence for robustbase
      if (inherits(model, "lmrob")) {
        doc <- doc %>%
          body_add_par(paste("Model converged:", model$converged), style = "Normal")
      }
      
      # Save the document
      print(doc, target = file)
    }
  )
}

shinyApp(ui = ui, server = server)
