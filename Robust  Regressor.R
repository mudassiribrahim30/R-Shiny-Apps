library(shiny)
library(haven)
library(readxl)
library(robustbase)
library(dplyr)
library(DT)
library(officer)
library(lm.beta)
library(tidyverse)

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
      
      /* Add notification animation */
      .update-notification {
        animation: slideIn 0.5s ease-out;
      }
      
      @keyframes slideIn {
        from {
          transform: translateX(100%);
          opacity: 0;
        }
        to {
          transform: translateX(0);
          opacity: 1;
        }
      }
      
      #close-notification:hover {
        background-color: #0056b3;
      }
    ")),  # This comma separates the style from the script
    
    tags$script(HTML("
      $(document).ready(function() {
        // Check if notification was already shown today
        var lastShown = localStorage.getItem('lastUpdateNotification');
        var today = new Date().toDateString();
        
        if (!lastShown || lastShown !== today) {
          // Show notification after 2 seconds
          setTimeout(function() {
            $('#update-notification').fadeIn();
          }, 2000);
          
          // Store today's date
          localStorage.setItem('lastUpdateNotification', today);
        }
        
        // Close button functionality
        $('#close-notification').click(function() {
          $('#update-notification').fadeOut();
        });
        
        // Auto-hide after 15 seconds
        setTimeout(function() {
          $('#update-notification').fadeOut();
        }, 15000);
      });
    "))
  ),
  
  div(class = "logo-container",
      img(src = "https://cdn-icons-png.flaticon.com/512/2103/2103633.png", 
          class = "logo-img", alt = "Robust Regressor Logo")
  ),
  
  title = "Robust Regressor",
  
  titlePanel(
    div(
      h1("Robust Regressor", style = "color: #2c3e50; text-align: center;"),
      p("A tool for robust regression analysis with categorical variable support", 
        style = "color: #7f8c8d; font-style: italic; text-align: center;")
    )
  ),
  
  # Update Notification Container
  tags$div(
    id = "update-notification",
    class = "shiny-notification",
    style = "position: fixed; top: 20px; right: 20px; width: 350px; 
           background-color: #f8f9fa; border-left: 4px solid #007bff;
           padding: 15px; box-shadow: 0 4px 8px rgba(0,0,0,0.1);
           z-index: 9999; display: none;",
    tags$h4(style = "margin-top: 0; color: #0056b3;", 
            icon("bell"), "Latest Updates"),
    tags$p(style = "margin: 10px 0; font-weight: bold;", 
           "Update and fixed issues: 21st January, 2026"),
    tags$ul(style = "padding-left: 20px; margin: 10px 0;",
            tags$li("Fixed R vs. Shiny discrepancies with reference categories."),
            tags$li("Fixed data loading issues with special characters"),
            tags$li("App now relies solely on robustbase for analysis"),
            tags$li("Improved Word report generation")
    ),
    tags$button(
      id = "close-notification",
      style = "background-color: #007bff; color: white; border: none;
             padding: 5px 15px; border-radius: 3px; cursor: pointer;",
      "Close"
    )
  ),
  
  sidebarLayout(
    sidebarPanel(
      div(class = "well-panel",
          h4("Data Input"),
          fileInput("file", "Upload Data File (Excel, CSV)",
                    accept = c(".csv", ".sav", ".dta", ".xlsx")),
      ),
      
      uiOutput("var_ui"),
      
      uiOutput("ref_ui"),  # Compact reference selection UI
      
      div(class = "well-panel",
          h4("Model Settings"),
          selectInput("method", "Select Method (lmrob)",
                      choices = c("MM", "M", "S"), selected = "MM",
                      width = "100%"),
          
          numericInput("decimal_places", "Decimal Places", 
                       value = 4, min = 1, max = 6, width = "100%")
      ),
      
      actionButton("run", "Run Robust Regression", 
                   class = "btn btn-primary",
                   style = "width: 100%; font-weight: bold;")
    ),
    
    mainPanel(
      tabsetPanel(
        tabPanel("Model Output", 
                 icon = icon("chart-line"),
                 div(class = "large-output", verbatimTextOutput("model_out"))),
        
        tabPanel("Estimates Table", 
                 icon = icon("table"),
                 DTOutput("estimates_table"),
                 uiOutput("estimates_interpretation")),
        
        tabPanel("Data View", 
                 icon = icon("database"),
                 h4("Data Preview"),
                 DTOutput("data_view")),
        
        tabPanel("Download Full Results", 
                 icon = icon("file-download"),
                 h4("Download Complete Analysis"),
                 downloadButton("download_report", "Download Word Report",
                                class = "btn btn-success")),
        
        tabPanel("About",
                 icon = icon("info-circle"),
                 div(class = "about-section",
                     h3("About Robust Regressor"),
                     p(HTML("Robust Regressor is a user-friendly Shiny application for robust regression analysis. 
       It directly uses the <span style='color: #e74c3c; font-weight: bold;'>robustbase</span> 
       package to provide reliable estimates that are resistant to outliers and data irregularities, 
       ensuring your statistical results are trustworthy."),
                       style = "color: #2c3e50; font-size: 16px; text-align: justify; line-height: 1.5; margin-top: 10px;"),
                     
                     
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
                     
                     h4("About the Methods"),
                     p("MM estimator: Recommended for most applications (balanced robustness and efficiency)"),
                     p("M estimator: More robust but less efficient"),
                     p("S estimator: Has high breakdown point"),
                     p("Updated and issues fixed: 21st January, 2026",
                       style = "background-color: #f9ed69; color: #000; 
           font-weight: bold; text-align: center; 
           padding: 10px; border-radius: 5px; 
           margin-top: 15px;"),
                     
                     
                     div(class = "references",
                         h4("References"),
                         p("Maechler, M., Rousseeuw, P., Croux, C., Todorov, V., Ruckstuhl, A., Salibian-Barrera, M., ... & Verbeke, T. (2021). robustbase: Basic Robust Statistics. R package version 0.93-9."),
                         p("Huber, P. J. (1981). Robust Statistics. Wiley."),
                         
                         h4("Developer Information"),
                         p("For suggestions or problems, please contact: mudassiribrahim30@gmail.com")
                     )
                 )
        )
      )
    )
  )
)

server <- function(input, output, session) {
  # Load data
  data <- reactive({
    req(input$file)
    ext <- tools::file_ext(input$file$name)
    
    df <- tryCatch({
      switch(ext,
             csv  = read.csv(input$file$datapath, stringsAsFactors = FALSE),
             sav  = read_sav(input$file$datapath),
             dta  = read_dta(input$file$datapath),
             xlsx = read_excel(input$file$datapath),
             validate("Unsupported file format")
      )
    }, error = function(e) {
      showNotification(paste("Error reading file:", e$message), type = "error")
      return(NULL)
    })
    
    if (is.null(df)) {
      validate("Failed to read the file. Please check the file format.")
    }
    
    # Clean column names - replace spaces with underscores
    names(df) <- gsub(" ", "_", names(df))
    names(df) <- gsub("[^[:alnum:]_\\.]", "_", names(df))
    
    df
  })
  
  # UI for selecting dependent and independent variables
  output$var_ui <- renderUI({
    req(data())
    vars <- names(data())
    tagList(
      div(class = "well-panel",
          h4("Variable Selection"),
          selectInput("y", "Dependent Variable", vars, width = "100%"),
          selectInput("x", "Independent Variables", vars, multiple = TRUE, width = "100%")
      )
    )
  })
  
  # Compact UI for changing reference levels of factor variables
  output$ref_ui <- renderUI({
    req(data(), input$x)
    df <- data()
    
    # Convert character columns to factors for selected independent variables
    df[input$x] <- lapply(df[input$x], function(col) {
      if (is.character(col)) as.factor(col) else col
    })
    
    # Identify factor independent variables
    factor_vars <- input$x[sapply(df[, input$x, drop = FALSE], is.factor)]
    if (length(factor_vars) == 0) return(NULL)
    
    # Create a single panel with all factor variables for reference selection
    div(class = "well-panel",
        h4("Reference Categories"),
        lapply(factor_vars, function(var) {
          div(class = "ref-cat-box",
              selectInput(
                paste0("ref_", var),
                label = var,
                choices = levels(df[[var]]),
                selected = levels(df[[var]])[1]  # default first level
              )
          )
        })
    )
  })
  
  # Process data with reference levels
  processed_data <- reactive({
    req(input$y, input$x, data())
    df <- data()[, c(input$y, input$x)]
    df <- na.omit(df)
    
    # Convert character columns to factors if any
    df[input$x] <- lapply(df[input$x], function(col) {
      if (is.character(col)) as.factor(col) else col
    })
    
    # Update reference levels for factors if user selected new ones
    for (var in input$x) {
      ref_input <- input[[paste0("ref_", var)]]
      if (!is.null(ref_input)) {
        df[[var]] <- relevel(df[[var]], ref = ref_input)
      }
    }
    
    df
  })
  
  # Run robust regression
  model <- eventReactive(input$run, {
    req(input$y, input$x, processed_data())
    df <- processed_data()
    
    # Build formula and run robust regression
    formula <- as.formula(
      paste(input$y, "~", paste(input$x, collapse = "+"))
    )
    
    tryCatch({
      lmrob(formula, data = df, method = input$method)
    }, error = function(e) {
      showNotification(paste("Error in model fitting:", e$message), type = "error")
      return(NULL)
    })
  })
  
  # Show model output
  output$model_out <- renderPrint({
    req(model())
    summary(model())
  })
  
  # Function to manually extract coefficients and create estimates table
  create_estimates_table <- function(model) {
    req(model)
    
    # Get model summary
    summ <- summary(model)
    
    # Extract coefficients table from summary
    coef_table <- summ$coefficients
    
    # Get model data for standardization
    model_data <- model$model
    
    # Create a temporary OLS model for standardization
    temp_lm <- tryCatch({
      lm(model$terms, data = model_data)
    }, error = function(e) {
      return(NULL)
    })
    
    # Get standardized coefficients if possible
    std_coefs <- if (!is.null(temp_lm)) {
      beta_result <- tryCatch({
        lm.beta(temp_lm)$standardized.coefficients
      }, error = function(e) {
        return(NULL)
      })
      
      if (!is.null(beta_result)) {
        beta_result
      } else {
        rep(NA, nrow(coef_table))
      }
    } else {
      rep(NA, nrow(coef_table))
    }
    
    # Create the estimates dataframe
    estimates_df <- data.frame(
      Term = rownames(coef_table),
      Estimate = coef_table[, "Estimate"],
      Std_Error = coef_table[, "Std. Error"],
      t_value = coef_table[, "t value"],
      p_value = coef_table[, "Pr(>|t|)"],
      stringsAsFactors = FALSE
    )
    
    # Add standardized coefficients
    estimates_df$Std_Estimate <- std_coefs[match(estimates_df$Term, names(std_coefs))]
    
    # Calculate confidence intervals
    estimates_df$Lower_CI <- estimates_df$Estimate - 1.96 * estimates_df$Std_Error
    estimates_df$Upper_CI <- estimates_df$Estimate + 1.96 * estimates_df$Std_Error
    
    # Add significance indicator
    estimates_df$Significance <- ifelse(
      estimates_df$p_value < 0.05 & !is.na(estimates_df$p_value),
      "Significant",
      ifelse(is.na(estimates_df$p_value), "NA", "Not Significant")
    )
    
    # Round numeric values
    decimal_places <- ifelse(is.null(input$decimal_places), 4, input$decimal_places)
    estimates_df <- estimates_df %>%
      mutate(across(where(is.numeric), ~round(.x, decimal_places)))
    
    return(estimates_df)
  }
  
  # Display estimates table
  output$estimates_table <- renderDT({
    req(model())
    
    estimates_df <- create_estimates_table(model())
    
    # Select and rename columns for display
    display_df <- estimates_df %>%
      select(
        Term,
        Estimate,
        Std_Estimate,
        Std_Error,
        Lower_CI,
        Upper_CI,
        p_value,
        Significance
      )
    
    # Create datatable
    dt <- DT::datatable(
      display_df,
      options = list(
        dom = 't',
        pageLength = nrow(display_df),
        scrollX = TRUE,
        autoWidth = TRUE
      ),
      rownames = FALSE,
      caption = "Robust Regression Estimates",
      colnames = c(
        'Term',
        'Unstandardized Estimate',
        'Standardized Estimate',
        'Std. Error',
        'Lower 95% CI',
        'Upper 95% CI',
        'p-value',
        'Significance'
      )
    ) %>%
      formatStyle(
        'Significance',
        backgroundColor = styleEqual(
          c('Significant', 'Not Significant', 'NA'),
          c('#e6ffe6', '#ffffff', '#f5f5f5')
        )
      )
    
    dt
  })
  
  # Interpretation for estimates table
  output$estimates_interpretation <- renderUI({
    req(model())
    
    div(class = "interpretation",
        h4("Interpreting the Estimates Table"),
        tags$ul(
          tags$li(tags$b("Unstandardized Estimate:"), " The change in the dependent variable for a one-unit change in the predictor, holding other variables constant"),
          tags$li(tags$b("Standardized Estimate:"), " Allows comparison of effect sizes across different variables (unitless)"),
          tags$li(tags$b("Std. Error:"), " The standard deviation of the sampling distribution of the estimate"),
          tags$li(tags$b("95% Confidence Interval:"), " The range within which the true population parameter is likely to fall"),
          tags$li(tags$b("p-value:"), " The probability of observing the result if the null hypothesis (no effect) is true"),
          tags$li(tags$b("Significance:"), " Variables with p < 0.05 (green highlight) are statistically significant at the 5% level")
        ),
        p("Note: These estimates are robust to outliers and influential observations, making them more reliable than OLS estimates when data contains anomalies.")
    )
  })
  
  # Data view tab
  output$data_view <- renderDT({
    req(data())
    DT::datatable(
      data(),
      options = list(
        pageLength = 10,
        scrollX = TRUE,
        dom = 'Bfrtip'
      ),
      class = 'cell-border stripe',
      caption = "Data Preview (First 10 rows shown)"
    )
  })
  
  # Download handler for Word report
  output$download_report <- downloadHandler(
    filename = function() {
      paste("robust_regression_report_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(model())
      
      # Get estimates table
      estimates_df <- create_estimates_table(model())
      
      # Create a new Word document
      doc <- read_docx() %>%
        body_add_par("ROBUST REGRESSION ANALYSIS REPORT", style = "heading 1") %>%
        body_add_par("") %>%
        body_add_par(paste("Generated by Robust Regressor on", Sys.Date()), style = "Normal") %>%
        body_add_par("") %>%
        
        # Add analysis details
        body_add_par("ANALYSIS DETAILS", style = "heading 2") %>%
        body_add_par("") %>%
        body_add_par(paste("Dependent Variable:", input$y), style = "Normal") %>%
        body_add_par(paste("Independent Variables:", paste(input$x, collapse = ", ")), style = "Normal") %>%
        body_add_par(paste("Method:", input$method), style = "Normal") %>%
        body_add_par(paste("Sample Size:", nrow(processed_data())), style = "Normal") %>%
        body_add_par("") %>%
        
        # Add reference categories if any
        body_add_par("REFERENCE CATEGORIES", style = "heading 2") %>%
        body_add_par("")
      
      # Add reference category information
      ref_cats <- list()
      for (var in input$x) {
        ref_input <- input[[paste0("ref_", var)]]
        if (!is.null(ref_input)) {
          ref_cats[[var]] <- ref_input
        }
      }
      
      if (length(ref_cats) > 0) {
        for (var in names(ref_cats)) {
          doc <- doc %>% 
            body_add_par(paste(var, "-> Reference:", ref_cats[[var]]), style = "Normal")
        }
      } else {
        doc <- doc %>% 
          body_add_par("No categorical variables with custom reference levels", style = "Normal")
      }
      
      doc <- doc %>% body_add_par("") %>%
        
        # Add model summary
        body_add_par("MODEL SUMMARY", style = "heading 2") %>%
        body_add_par("")
      
      # Capture model summary
      model_summary <- capture.output(summary(model()))
      for (line in model_summary) {
        doc <- doc %>% body_add_par(line, style = "Normal")
      }
      
      doc <- doc %>% body_add_par("") %>%
        
        # Add detailed estimates table
        body_add_par("DETAILED ESTIMATES TABLE", style = "heading 2") %>%
        body_add_par("")
      
      # Prepare table for Word document
      table_df <- estimates_df %>%
        select(
          Term,
          Estimate,
          Std_Estimate,
          Std_Error,
          Lower_CI,
          Upper_CI,
          p_value
        ) %>%
        rename(
          Predictor = Term,
          b = Estimate,
          β = Std_Estimate,
          SE = Std_Error,
          `Lower CI` = Lower_CI,
          `Upper CI` = Upper_CI,
          `p-value` = p_value
        )
      
      # Add the table
      doc <- doc %>% 
        body_add_table(
          table_df,
          header = TRUE,
          style = "table_template"
        ) %>%
        body_add_par("") %>%
        
        # Add interpretation
        body_add_par("INTERPRETATION", style = "heading 2") %>%
        body_add_par("") %>%
        body_add_par("Key Findings:", style = "Normal") %>%
        body_add_par("")
      
      # Add significant variables
      sig_vars <- estimates_df %>%
        filter(Significance == "Significant") %>%
        pull(Term)
      
      if (length(sig_vars) > 0) {
        doc <- doc %>% 
          body_add_par("Statistically Significant Variables:", style = "Normal")
        for (var in sig_vars) {
          var_info <- estimates_df %>% filter(Term == var)
          doc <- doc %>% 
            body_add_par(paste("•", var, 
                               "(b =", round(var_info$Estimate, 3),
                               ", p =", format.pval(var_info$p_value, digits = 3), ")"), 
                         style = "Normal")
        }
      } else {
        doc <- doc %>% 
          body_add_par("No statistically significant variables at p < 0.05 level", 
                       style = "Normal")
      }
      
      doc <- doc %>% 
        body_add_par("") %>%
        body_add_par("Notes:", style = "Normal") %>%
        body_add_par("1. b = unstandardized coefficient", style = "Normal") %>%
        body_add_par("2. β = standardized coefficient", style = "Normal") %>%
        body_add_par("3. All estimates are robust to outliers and influential points", style = "Normal") %>%
        body_add_par("4. Confidence intervals are approximate 95% intervals", style = "Normal")
      
      # Save the document
      print(doc, target = file)
    }
  )
}

shinyApp(ui = ui, server = server)
