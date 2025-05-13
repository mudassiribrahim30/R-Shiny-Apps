library(shiny)
library(lavaan)
library(semPlot)
library(ggplot2)
library(readr)
library(readxl)
library(DT)
library(officer)
library(flextable)
library(shinythemes)

app_name <- "MedModr"
app_version <- "1.1.0"

ui <- fluidPage(
  theme = shinytheme("cosmo"),
  titlePanel(
    div(
      img(src = "https://i.imgur.com/xyZQl9E.png", height = 60, style = "margin-right:15px;"),
      span(app_name, style = "vertical-align:middle; font-weight:bold; font-size:28px;")
    )
  ),
  
  # Custom CSS
  tags$head(
    tags$style(HTML("
      #fitText, #summaryText, #effectsText {
        font-size: 16px !important;
        line-height: 1.5 !important;
      }
      .shiny-html-output {
        font-size: 16px !important;
      }
      .well {
        background-color: #f9f9f9;
        border-radius: 8px;
      }
      .btn-primary {
        background-color: #428bca;
        border-color: #357ebd;
      }
      .btn-primary:hover {
        background-color: #3276b1;
        border-color: #285e8e;
      }
      .tab-content {
        padding: 15px;
        border-left: 1px solid #ddd;
        border-right: 1px solid #ddd;
        border-bottom: 1px solid #ddd;
        border-radius: 0 0 4px 4px;
      }
    "))
  ),
  
  sidebarLayout(
    sidebarPanel(
      width = 4,
      h4("Data Input", icon("database")),
      fileInput("datafile", "Upload your data (CSV, XLSX, or XLS):", 
                accept = c(".csv", ".xlsx", ".xls")),
      
      uiOutput("var_preview"),
      
      h4("Analysis Settings", icon("cogs")),
      selectInput("analysis_type", "Select Analysis Type:",
                  choices = c("Simple Mediation", "Serial Mediation", "Moderation")),
      
      checkboxInput("use_composites", "Use Constructs (latent variables with indicators)?", value = FALSE),
      
      numericInput("bootstrap_samples", "Number of Bootstrap Samples:", 
                   value = 1000, min = 100, max = 10000, step = 100),
      
      checkboxInput("use_bootstrap", "Use Bootstrap Confidence Intervals?", value = FALSE),
      
      conditionalPanel(
        condition = "input.use_bootstrap == true",
        helpText("Note: Bootstrapping will take longer but provides more robust results.")
      ),
      
      conditionalPanel(
        condition = "input.analysis_type == 'Simple Mediation'",
        h4("Simple Mediation Guide:", icon("info-circle")),
        tags$ul(
          tags$li("X = Independent Variable"),
          tags$li("M = Mediator Variable"),
          tags$li("Y = Dependent Variable"),
          tags$li("a = X → M path"),
          tags$li("b = M → Y path (controlling for X)"),
          tags$li("c' = Direct effect of X → Y (controlling for M)")
        ),
        h4("Example Syntax:"),
        verbatimTextOutput("simple_mediation_syntax")
      ),
      
      conditionalPanel(
        condition = "input.analysis_type == 'Serial Mediation'",
        h4("Serial Mediation Guide:", icon("info-circle")),
        tags$ul(
          tags$li("X = Independent Variable"),
          tags$li("M1 = Mediator 1"),
          tags$li("M2 = Mediator 2"),
          tags$li("Y = Dependent Variable"),
          tags$li("a1 = X → M1 path"),
          tags$li("a2 = M1 → M2 path"),
          tags$li("b = M2 → Y path"),
          tags$li("c' = Direct effect of X → Y")
        ),
        h4("Example Syntax:"),
        verbatimTextOutput("serial_mediation_syntax"),
        h4("Important Notes:"),
        tags$ul(
          tags$li("All mediators must be specified in order"),
          tags$li("Make sure your model syntax matches the serial mediation structure"),
          tags$li("Check that all variable names match your data")
        )
      ),
      
      conditionalPanel(
        condition = "input.analysis_type == 'Moderation'",
        h4("Moderation Guide:", icon("info-circle")),
        tags$ul(
          tags$li("X = Independent Variable"),
          tags$li("W = Moderator Variable"),
          tags$li("Y = Dependent Variable"),
          tags$li("XW = Interaction Term (X × W)"),
          tags$li("Interpret significant interaction with simple slopes analysis")
        ),
        h4("Example Syntax:"),
        verbatimTextOutput("moderation_syntax"),
        h4("Important Notes:"),
        tags$ul(
          tags$li("You must create the interaction term in your data first"),
          tags$li("Name it something like 'XW' or 'X_W'"),
          tags$li("Center your variables before creating interaction terms")
        )
      ),
      
      h4("Model Specification", icon("code")),
      textAreaInput("model_syntax", "Specify lavaan Model Syntax:", 
                    placeholder = "Enter your model syntax here...", height = "200px"),
      
      actionButton("run", "Run Analysis", icon = icon("play"), 
                   class = "btn-primary btn-lg"),
      br(), br(),
      downloadButton("download_report", "Download Word Report", class = "btn-success"),
      downloadButton("download_plot", "Download Diagram (PNG)", class = "btn-info")
    ),
    
    mainPanel(
      width = 8,
      tabsetPanel(
        tabPanel("Diagram", 
                 plotOutput("semPlot", height = "600px"),
                 h4("Diagram Interpretation:"),
                 tags$ul(
                   tags$li("Arrows represent hypothesized relationships between variables"),
                   tags$li("Standardized coefficients are shown on the paths"),
                   tags$li("Solid lines indicate significant relationships (p < .05)"),
                   tags$li("Dashed lines indicate non-significant paths"),
                   tags$li("Rectangles represent observed variables"),
                   tags$li("Ovals represent latent variables (if used)")
                 )),
        tabPanel("Results", 
                 h3("Model Fit Indices", icon("check-circle")),
                 verbatimTextOutput("fitText"),
                 h3("Parameter Estimates", icon("table")),
                 verbatimTextOutput("summaryText"),
                 h3("Effects Analysis", icon("project-diagram")),
                 verbatimTextOutput("effectsText"),
                 h4("Results Interpretation:"),
                 tags$ul(
                   tags$li("Check model fit indices first (CFI > .90, RMSEA < .08 indicate good fit)"),
                   tags$li(paste("Bootstrap CIs are", 
                                 textOutput("bootstrap_status", inline = TRUE))),
                   tags$li("For mediation: Significant indirect effect (a*b) indicates mediation"),
                   tags$li("For serial mediation: Check all indirect paths (a1*a2*b)"),
                   tags$li("For moderation: Significant interaction term (X:W) indicates moderation"),
                   tags$li("p-values < .05 are typically considered statistically significant"),
                   tags$li("Standardized coefficients (std.all) show effect sizes")
                 )),
        tabPanel("Variables in Data", 
                 DTOutput("var_table"),
                 br(),
                 h4("Data Summary"),
                 verbatimTextOutput("data_summary")),
        tabPanel("About",
                 div(class = "well",
                     h4(app_name, "v", app_version),
                     p("A user-friendly interface for mediation, serial mediation, and moderation analysis using composite variables or constructs."),
                     h5("Key Features:"),
                     tags$ul(
                       tags$li("Supports simple and serial mediation models"),
                       tags$li("Moderation analysis with interaction terms"),
                       tags$li("Bootstrap confidence intervals"),
                       tags$li("Interactive path diagrams"),
                       tags$li("Professional report generation")
                     ),
                     h5("Developed by:"),
                     p("Mudasir Mohammed Ibrahim"),
                     h5("Contact:"),
                     p("mudassiribrahim30@gmail.com"),
                     h5("License:"),
                     p("MIT License - Free for academic and research use"),
                     h5("Citation:"),
                     p("Please cite this app if used in your research publications")
                 ))
      )
    )
  )
)

server <- function(input, output, session) {
  
  # Read data
  data <- reactive({
    req(input$datafile)
    ext <- tools::file_ext(input$datafile$name)
    tryCatch({
      if (ext == "csv") {
        read_csv(input$datafile$datapath)
      } else {
        read_excel(input$datafile$datapath)
      }
    }, error = function(e) {
      showNotification(paste("Error reading file:", e$message), type = "error", duration = 10)
      return(NULL)
    })
  })
  
  # Variable preview
  output$var_preview <- renderUI({
    req(data())
    var_names <- names(data())
    if(length(var_names) > 15) {
      var_text <- paste0(paste(var_names[1:15], collapse = ", "), ", ... (+", length(var_names)-15, " more)")
    } else {
      var_text <- paste(var_names, collapse = ", ")
    }
    HTML(paste0("<strong>Variables in dataset (", ncol(data()), "):</strong><br>", var_text))
  })
  
  output$var_table <- renderDT({
    req(data())
    datatable(data(), 
              options = list(
                scrollX = TRUE,
                pageLength = 10,
                lengthMenu = c(5, 10, 15, 20)
              ))
  })
  
  output$data_summary <- renderPrint({
    req(data())
    df <- data()
    cat("DATA SUMMARY\n")
    cat("------------\n")
    cat("Number of observations:", nrow(df), "\n")
    cat("Number of variables:", ncol(df), "\n\n")
    
    cat("VARIABLE TYPES:\n")
    var_types <- sapply(df, function(x) class(x)[1])
    print(table(var_types))
    
    cat("\nMISSING VALUES:\n")
    missing_count <- sapply(df, function(x) sum(is.na(x)))
    if(any(missing_count > 0)) {
      print(missing_count[missing_count > 0])
      cat("\nNote: The analysis will automatically remove rows with missing values.\n")
    } else {
      cat("No missing values detected.\n")
    }
  })
  
  # Bootstrap status text
  output$bootstrap_status <- renderText({
    if(input$use_bootstrap) {
      paste("enabled with", input$bootstrap_samples, "samples")
    } else {
      "not used (using model-based standard errors)"
    }
  })
  
  # Syntax examples
  output$simple_mediation_syntax <- renderText({
    "M ~ a*X\nY ~ b*M + cp*X\n\n# Define effects\nindirect := a*b\ntotal := cp + (a*b)\n# Proportion mediated\nprop_mediated := (a*b)/total"
  })
  
  output$serial_mediation_syntax <- renderText({
    "M1 ~ a1*X\nM2 ~ a2*M1 + d*X\nY ~ b*M2 + cp*X\n\n# Define effects\nindirect1 := a1*a2*b  # via M1 and M2\nindirect2 := a1*d*b   # via M1 only\ntotal := cp + (a1*a2*b) + (a1*d*b)"
  })
  
  output$moderation_syntax <- renderText({
    "# Note: You must create the interaction term in your data first\n# (e.g., data$XW <- data$X * data$W)\n\nY ~ b1*X + b2*W + b3*XW\n\n# Define simple slopes\nlowW := b1 + b3*(-1)\navgW := b1 + b3*(0)\nhighW := b1 + b3*(1)"
  })
  
  # Model estimation
  model_fit <- eventReactive(input$run, {
    req(input$model_syntax, data())
    df <- na.omit(data())
    
    # For moderation analysis, check if interaction term exists
    if(input$analysis_type == "Moderation") {
      vars <- all.vars(parse(text = input$model_syntax))
      if(any(grepl("XW|X_W|interaction", vars, ignore.case = TRUE)) && 
         !any(grepl("XW|X_W|interaction", names(df), ignore.case = TRUE))) {
        showNotification("Error: Interaction term not found in data. Please create it first.", 
                         type = "error", duration = 10)
        return(NULL)
      }
    }
    
    # For serial mediation, validate model structure
    if(input$analysis_type == "Serial Mediation") {
      if(!grepl("M1.*X.*M2.*M1.*Y.*M2", input$model_syntax)) {
        showNotification("Warning: Serial mediation model should include paths from X to M1, M1 to M2, and M2 to Y", 
                         type = "warning", duration = 10)
      }
    }
    
    tryCatch({
      if(input$use_bootstrap) {
        # Run with bootstrap
        sem(model = input$model_syntax, 
            data = df, 
            std.lv = TRUE,
            se = "bootstrap", 
            bootstrap = input$bootstrap_samples)
      } else {
        # Run without bootstrap
        sem(model = input$model_syntax, 
            data = df, 
            std.lv = TRUE)
      }
    }, error = function(e) {
      showNotification(paste("Error in model estimation:", e$message), 
                       type = "error", duration = 10)
      return(NULL)
    })
  })
  
  output$fitText <- renderPrint({
    req(model_fit())
    fit <- fitMeasures(model_fit(), c("chisq", "df", "pvalue", "cfi", "tli", "rmsea", "rmsea.ci.lower", "rmsea.ci.upper", "srmr"))
    cat("MODEL FIT INDICES:\n")
    cat(sprintf("Chi-square(%.0f) = %.3f, p = %.3f\n", fit["df"], fit["chisq"], fit["pvalue"]))
    cat(sprintf("CFI = %.3f, TLI = %.3f\n", fit["cfi"], fit["tli"]))
    cat(sprintf("RMSEA = %.3f, 90%% CI [%.3f, %.3f]\n", fit["rmsea"], fit["rmsea.ci.lower"], fit["rmsea.ci.upper"]))
    cat(sprintf("SRMR = %.3f\n\n", fit["srmr"]))
    
    cat("Fit Interpretation:\n")
    if (fit["cfi"] > 0.95 && fit["rmsea"] < 0.06) {
      cat("Excellent model fit (CFI > 0.95, RMSEA < 0.06)\n")
    } else if (fit["cfi"] > 0.90 && fit["rmsea"] < 0.08) {
      cat("Acceptable model fit (CFI > 0.90, RMSEA < 0.08)\n")
    } else {
      cat("Poor model fit - consider revising your model\n")
    }
    
    if(input$use_bootstrap) {
      cat(sprintf("\nBootstrap Results (based on %d samples):\n", input$bootstrap_samples))
    }
  })
  
  output$summaryText <- renderPrint({
    req(model_fit())
    
    if(input$use_bootstrap) {
      pe <- parameterEstimates(model_fit(), 
                               standardized = TRUE, 
                               ci = TRUE, 
                               boot.ci.type = "perc",
                               level = 0.95)
    } else {
      pe <- parameterEstimates(model_fit(), 
                               standardized = TRUE, 
                               ci = TRUE)
    }
    
    pe <- pe[pe$op != ":=", ]  # Exclude defined parameters
    
    cat("STANDARDIZED PARAMETER ESTIMATES:\n\n")
    print(pe[, c("lhs", "op", "rhs", "est", "se", "z", "pvalue", "ci.lower", "ci.upper", "std.all")], 
          row.names = FALSE)
    
    cat("\nInterpretation:\n")
    cat("- 'est' = unstandardized coefficient\n")
    cat("- 'std.all' = standardized coefficient (effect size)\n")
    cat("- p < .05 indicates statistical significance\n")
    cat("- 95% CIs that don't include 0 indicate significance\n")
    if(input$use_bootstrap) {
      cat("- Standard errors and CIs are based on bootstrap\n")
    } else {
      cat("- Standard errors and CIs are model-based\n")
    }
  })
  
  output$effectsText <- renderPrint({
    req(model_fit())
    
    if(input$use_bootstrap) {
      pe <- parameterEstimates(model_fit(), 
                               standardized = TRUE, 
                               ci = TRUE, 
                               boot.ci.type = "perc",
                               level = 0.95)
    } else {
      pe <- parameterEstimates(model_fit(), 
                               standardized = TRUE, 
                               ci = TRUE)
    }
    
    # Direct effects
    direct <- subset(pe, op == "~")
    if (nrow(direct) > 0) {
      cat("DIRECT EFFECTS:\n\n")
      print(direct[, c("lhs", "op", "rhs", "est", "se", "pvalue", "std.all")], 
            row.names = FALSE)
    }
    
    # Defined effects (indirect, total, etc.)
    defined <- subset(pe, op == ":=")
    if (nrow(defined) > 0) {
      cat("\nDEFINED EFFECTS:\n\n")
      print(defined[, c("lhs", "op", "rhs", "est", "se", "pvalue", "ci.lower", "ci.upper")], 
            row.names = FALSE)
      
      # Interpretation for mediation
      if (any(grepl("indirect", defined$lhs))) {
        indirect <- defined[grepl("indirect", defined$lhs), ]
        if (nrow(indirect) > 0 && indirect$pvalue[1] < 0.05) {
          cat(sprintf("\n- Significant indirect effect (%.3f, p = %.3f)\n", indirect$est[1], indirect$pvalue[1]))
          if ("total" %in% defined$lhs) {
            total <- defined[defined$lhs == "total", "est"]
            prop <- indirect$est[1]/total
            cat(sprintf("- Proportion mediated: %.1f%%\n", prop*100))
          }
        } else if (nrow(indirect) > 0) {
          cat("\n- Non-significant indirect effect\n")
        }
      }
      
      # Interpretation for serial mediation
      if (input$analysis_type == "Serial Mediation" && any(grepl("indirect1|indirect2", defined$lhs))) {
        if ("indirect1" %in% defined$lhs) {
          indirect1 <- defined[defined$lhs == "indirect1", ]
          cat(sprintf("\nSerial Mediation via M1 and M2: %.3f (p = %.3f)\n", 
                      indirect1$est, indirect1$pvalue))
        }
        if ("indirect2" %in% defined$lhs) {
          indirect2 <- defined[defined$lhs == "indirect2", ]
          cat(sprintf("Serial Mediation via M1 only: %.3f (p = %.3f)\n", 
                      indirect2$est, indirect2$pvalue))
        }
        if ("total" %in% defined$lhs) {
          total <- defined[defined$lhs == "total", ]
          cat(sprintf("Total effect: %.3f (p = %.3f)\n", total$est, total$pvalue))
        }
      }
      
      # Interpretation for moderation
      if (any(grepl("lowW|avgW|highW", defined$lhs))) {
        cat("\nModeration Interpretation (Simple Slopes):\n")
        if ("lowW" %in% defined$lhs) {
          low <- defined[defined$lhs == "lowW", ]
          cat(sprintf("- Effect at low W (-1 SD): %.3f, p = %.3f\n", low$est, low$pvalue))
        }
        if ("avgW" %in% defined$lhs) {
          avg <- defined[defined$lhs == "avgW", ]
          cat(sprintf("- Effect at mean W: %.3f, p = %.3f\n", avg$est, avg$pvalue))
        }
        if ("highW" %in% defined$lhs) {
          high <- defined[defined$lhs == "highW", ]
          cat(sprintf("- Effect at high W (+1 SD): %.3f, p = %.3f\n", high$est, high$pvalue))
        }
      }
    } else {
      cat("\nNo additional effects defined in the model.\n")
      cat("Tip: Define effects using ':=' syntax (e.g., 'indirect := a*b')\n")
    }
    
    if(input$use_bootstrap) {
      cat("\nNote: Effects estimates use bootstrap standard errors and confidence intervals\n")
    }
  })
  
  output$semPlot <- renderPlot({
    req(model_fit())
    semPaths(model_fit(), 
             what = "std", 
             layout = "tree", 
             style = "lisrel",
             residuals = FALSE, 
             edge.label.cex = 1.2, 
             sizeMan = 8,
             sizeLat = 10,
             color = list(lat = "skyblue", man = "lightyellow"),
             edge.color = "black",
             node.width = 1.5, 
             node.height = 1.5,
             fade = FALSE,
             edge.label.position = 0.6,
             rotation = 2)
  })
  
  output$download_report <- downloadHandler(
    filename = function() paste0("MedModr_Report_", Sys.Date(), ".docx"),
    content = function(file) {
      fit <- model_fit()
      if(is.null(fit)) {
        showNotification("Cannot generate report - no valid model", type = "error")
        return()
      }
      
      # Model fit
      fit_measures <- fitMeasures(fit, c("chisq", "df", "pvalue", "cfi", "tli", "rmsea", "srmr"))
      fit_text <- sprintf("Chi-square(%.0f) = %.3f, p = %.3f\nCFI = %.3f\nTLI = %.3f\nRMSEA = %.3f\nSRMR = %.3f",
                          fit_measures["df"], fit_measures["chisq"], fit_measures["pvalue"],
                          fit_measures["cfi"], fit_measures["tli"],
                          fit_measures["rmsea"], fit_measures["srmr"])
      
      # Parameter estimates
      if(input$use_bootstrap) {
        pe <- parameterEstimates(fit, standardized = TRUE, ci = TRUE, boot.ci.type = "perc")
      } else {
        pe <- parameterEstimates(fit, standardized = TRUE, ci = TRUE)
      }
      
      direct <- subset(pe, op == "~")
      direct_text <- capture.output(print(direct[, c("lhs", "op", "rhs", "est", "se", "pvalue", "std.all")], 
                                          row.names = FALSE))
      
      # Effects
      defined <- subset(pe, op == ":=")
      effects_text <- if (nrow(defined) > 0) {
        capture.output(print(defined[, c("lhs", "op", "rhs", "est", "se", "pvalue", "ci.lower", "ci.upper")], 
                             row.names = FALSE))
      } else {
        "No additional effects defined in the model."
      }
      
      # Create plot for report
      plot_file <- tempfile(fileext = ".png")
      png(plot_file, width = 800, height = 600)
      semPaths(fit, what = "std", layout = "tree", style = "lisrel",
               residuals = FALSE, edge.label.cex = 1.0, sizeMan = 6, sizeLat = 8,
               color = list(lat = "skyblue", man = "lightyellow"),
               edge.color = "black", fade = FALSE)
      dev.off()
      
      doc <- read_docx() %>%
        body_add_par(app_name, style = "heading 1") %>%
        body_add_par("Analysis Results", style = "heading 2") %>%
        
        body_add_par("Analysis Settings", style = "heading 3") %>%
        body_add_par(sprintf("Analysis type: %s", input$analysis_type), style = "Normal") %>%
        body_add_par(sprintf("Bootstrap samples: %d", ifelse(input$use_bootstrap, input$bootstrap_samples, 0)), style = "Normal") %>%
        
        body_add_par("Model Fit Indices", style = "heading 3") %>%
        body_add_par(fit_text, style = "Normal") %>%
        
        body_add_par("Parameter Estimates", style = "heading 3") %>%
        body_add_par(paste(direct_text, collapse = "\n"), style = "Normal") %>%
        
        body_add_par("Defined Effects", style = "heading 3") %>%
        body_add_par(paste(effects_text, collapse = "\n"), style = "Normal") %>%
        
        body_add_par("Path Diagram", style = "heading 3") %>%
        body_add_img(plot_file, width = 6, height = 4.5) %>%
        
        body_add_par("Interpretation Notes", style = "heading 3") %>%
        body_add_par(sprintf("1. Bootstrap CIs %s", 
                             ifelse(input$use_bootstrap, 
                                    sprintf("enabled (%d samples)", input$bootstrap_samples),
                                    "not used")), 
                     style = "Normal") %>%
        body_add_par("2. Check model fit indices first (CFI > .90, RMSEA < .08 indicate acceptable fit)", style = "Normal") %>%
        body_add_par("3. Significant indirect effects (p < .05) indicate mediation", style = "Normal") %>%
        body_add_par("4. Standardized coefficients (std.all) show effect sizes", style = "Normal") %>%
        
        body_add_par("Generated by MedModr", style = "Normal") %>%
        body_add_par(paste("Date:", Sys.Date()), style = "Normal")
      
      print(doc, target = file)
    }
  )
  
  output$download_plot <- downloadHandler(
    filename = function() paste0("MedModr_Diagram_", Sys.Date(), ".png"),
    content = function(file) {
      req(model_fit())
      png(file, width = 1000, height = 800)
      semPaths(model_fit(), 
               what = "std", 
               layout = "tree", 
               style = "lisrel",
               residuals = FALSE, 
               edge.label.cex = 1.2, 
               sizeMan = 8, 
               sizeLat = 10,
               color = list(lat = "skyblue", man = "lightyellow"),
               edge.color = "black",
               fade = FALSE)
      dev.off()
    }
  )
}

shinyApp(ui, server)
