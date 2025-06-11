library(shiny)
library(shinythemes)
library(officer)
library(flextable)
library(magrittr)
library(pwr)
library(e1071)
library(tidyverse)

counter_file <- "counter.rds"
if (!file.exists(counter_file)) saveRDS(0, counter_file)
counter <- readRDS(counter_file) + 1
saveRDS(counter, counter_file)

getGreeting <- function() {
  hour <- as.numeric(format(Sys.time(), "%H"))
  if (hour < 12) {
    return("Good morning!")
  } else if (hour < 18) {
    return("Good afternoon!")
  } else {
    return("Good evening!")
  }
}

ui <- navbarPage(
  title = div(icon("chart-bar"), "CalcuStats"),
  theme = shinytheme("cosmo"),
  footer = div(
    style = "text-align: center; padding: 10px; background-color: #f8f9fa; border-top: 1px solid #ddd;",
    HTML("<p>For suggestions or support, contact <a href='mailto:mudassiribrahim30@gmail.com'>mudassiribrahim30@gmail.com</a></p>"),
    div(
      style = "margin-top: 10px; font-size: 0.9em; color: #555;",
      "Developed by Mudasir Mohammed Ibrahim (Registered Nurse, BSc, Dip)"
    )
  ),
  
  tabPanel("Proportional Allocation",
           sidebarLayout(
             sidebarPanel(
               width = 4,
               numericInput("custom_sample", "Your Sample Size (n)", value = 100, min = 1, step = 1),
               numericInput("non_response_custom", "Non-response Rate (%)", value = 0, min = 0, max = 100, step = 1),
               actionButton("addStratum_custom", "Add Stratum", class = "btn-primary"),
               actionButton("removeStratum_custom", "Remove Last Stratum", class = "btn-warning"),
               uiOutput("stratumInputs_custom"),
               div(
                 style = "margin-top: 10px; padding: 10px; background-color: #f8f9fa; border-radius: 5px;",
                 p("For a single population, enter the total size below."),
                 p("For multiple populations, click 'Add Stratum' to create additional groups.")
               ),
               downloadButton("downloadCustomWord", "Download as Word", class = "btn-success"),
               downloadButton("downloadCustomSteps", "Download Calculation Steps", class = "btn-info")
             ),
             mainPanel(
               width = 8,
               h3("Proportional Allocation for Specified Sample Size"),
               h4(textOutput("greetingTextCustom"), style = "color: #1A5276; font-weight: bold;"),
               div(
                 style = "padding: 10px; background-color: #f0f8ff; border-left: 5px solid #1A5276; margin-bottom: 20px;",
                 HTML("
            <strong>Proportional Allocation Formula:</strong><br>
            <em>n<sub>i</sub> = (N<sub>i</sub> / N) &times; n</em><br>
            <ul>
              <li><strong>n<sub>i</sub></strong> = sample size for stratum i</li>
              <li><strong>N<sub>i</sub></strong> = population size for stratum i</li>
              <li><strong>N</strong> = total population size</li>
              <li><strong>n</strong> = total sample size</li>
            </ul>
            This method distributes your specified sample size proportionally across strata based on their population sizes.
          ")
               ),
               fluidRow(
                 column(6, verbatimTextOutput("totalPopCustom")),
                 column(6, verbatimTextOutput("sampleSizeCustom"))
               ),
               verbatimTextOutput("adjustedSampleSizeCustom"),
               h4("Proportional Allocation"),
               tableOutput("allocationTableCustom"),
               textOutput("interpretationTextCustom")
             )
           )
  ),
  
  tabPanel("Taro Yamane",
           sidebarLayout(
             sidebarPanel(
               width = 4,
               numericInput("e", "Margin of Error (e)", value = 0.05, min = 0.0001, max = 1, step = 0.01),
               numericInput("non_response", "Non-response Rate (%)", value = 0, min = 0, max = 100, step = 1),
               actionButton("addStratum", "Add Stratum", class = "btn-primary"),
               actionButton("removeStratum", "Remove Last Stratum", class = "btn-warning"),
               uiOutput("stratumInputs"),
               div(
                 style = "margin-top: 10px; padding: 10px; background-color: #f8f9fa; border-radius: 5px;",
                 p("For a single population, enter the total size below."),
                 p("For multiple populations, click 'Add Stratum' to create additional groups.")
               ),
               downloadButton("downloadWord", "Download as Word", class = "btn-success"),
               downloadButton("downloadSteps", "Download Calculation Steps", class = "btn-info")
             ),
             mainPanel(
               width = 8,
               h3("Taro Yamane Method with Proportional Allocation"),
               h4(textOutput("greetingText"), style = "color: #1A5276; font-weight: bold;"),
               div(
                 style = "padding: 10px; background-color: #f0f8ff; border-left: 5px solid #1A5276; margin-bottom: 20px;",
                 HTML("
            <strong>Taro Yamane Formula:</strong><br>
            <em>n = N / (1 + N &times; e²)</em><br>
            <ul>
              <li><strong>n</strong> = required sample size</li>
              <li><strong>N</strong> = population size</li>
              <li><strong>e</strong> = margin of error (in decimal)</li>
            </ul>
            This formula assumes a 95% confidence level and simplifies sample size estimation for large populations.
          ")
               ),
               fluidRow(
                 column(6, verbatimTextOutput("totalPop")),
                 column(6, verbatimTextOutput("sampleSize"))
               ),
               verbatimTextOutput("formulaExplanation"),
               verbatimTextOutput("adjustedSampleSize"),
               h4("Proportional Allocation"),
               p("The proportional allocation formula is:"),
               p(HTML("<em>n<sub>i</sub> = (N<sub>i</sub> / N) &times; n</em>")),
               tableOutput("allocationTable"),
               textOutput("interpretationText")
             )
           )
  ),
  
  tabPanel("Cochran Formula",
           sidebarLayout(
             sidebarPanel(
               width = 4,
               numericInput("p", "Estimated Proportion (p)", value = 0.5, min = 0.01, max = 0.99, step = 0.01),
               numericInput("z", "Z-score (Z)", value = 1.96),
               numericInput("e_c", "Margin of Error (e)", value = 0.05, min = 0.0001, max = 1, step = 0.01),
               numericInput("N_c", "Population Size (optional)", value = NULL, min = 1),
               numericInput("non_response_c", "Non-response Rate (%)", value = 0, min = 0, max = 100, step = 1),
               actionButton("addStratum_c", "Add Stratum", class = "btn-primary"),
               actionButton("removeStratum_c", "Remove Last Stratum", class = "btn-warning"),
               uiOutput("stratumInputs_c"),
               downloadButton("downloadCochranWord", "Download as Word", class = "btn-success"),
               downloadButton("downloadCochranSteps", "Download Calculation Steps", class = "btn-info")
             ),
             mainPanel(
               width = 8,
               h3("Cochran Sample Size Calculator with Proportional Allocation"),
               div(
                 style = "padding: 10px; background-color: #f0f8ff; border-left: 5px solid #1A5276; margin-bottom: 20px;",
                 HTML("
            <strong>Cochran's Formula:</strong><br>
            <em>n₀ = (Z² × p × (1-p)) / e²</em><br>
            <ul>
              <li><strong>n₀</strong> = required sample size</li>
              <li><strong>Z</strong> = Z-score for desired confidence level</li>
              <li><strong>p</strong> = estimated proportion of the attribute present in the population</li>
              <li><strong>e</strong> = margin of error (in decimal)</li>
            </ul>
            For finite populations, apply the finite population correction: n = n₀ / (1 + (n₀ - 1)/N)
          ")
               ),
               verbatimTextOutput("cochranFormulaExplanation"),
               verbatimTextOutput("cochranSample"),
               verbatimTextOutput("cochranAdjustedSample"),
               h4("Proportional Allocation"),
               p("The proportional allocation formula is:"),
               p(HTML("<em>n<sub>i</sub> = (N<sub>i</sub> / N) &times; n</em>")),
               tableOutput("cochranAllocationTable"),
               textOutput("cochranInterpretation")
             )
           )
  ),
  
  tabPanel("Power Analysis",
           sidebarLayout(
             sidebarPanel(
               width = 4,
               selectInput("testType", "Statistical Test:",
                           choices = c("Independent t-test", "Paired t-test", "One-sample t-test", "One-Way ANOVA",
                                       "Two-Way ANOVA", "Proportion", "Correlation", "Chi-squared",
                                       "Simple Linear Regression", "Multiple Linear Regression")),
               numericInput("effectSize", "Effect Size", value = 0.5),
               numericInput("alpha", "Significance Level (alpha)", value = 0.05),
               numericInput("power", "Desired Power", value = 0.8),
               conditionalPanel(
                 condition = "input.testType == 'Multiple Linear Regression'",
                 numericInput("predictors", "Number of Predictors", value = 2, min = 1)
               ),
               actionButton("runPower", "Run Power Analysis", class = "btn-primary"),
               downloadButton("downloadPowerSteps", "Download Calculation Steps", class = "btn-info")
             ),
             mainPanel(
               width = 8,
               h3("Power Analysis for Inferential Tests"),
               verbatimTextOutput("powerResult")
             )
           )
  ),
  
  tabPanel("Descriptive Statistics",
           sidebarLayout(
             sidebarPanel(
               width = 4,
               tags$textarea(id = "dataInput", rows = 10, cols = 30,
                             placeholder = "Paste a column of data (with header) from Excel or statistical software..."),
               actionButton("runDesc", "Get Descriptive Statistics", class = "btn-primary"),
               downloadButton("downloadDescSteps", "Download Calculation Steps", class = "btn-info")
             ),
             mainPanel(
               width = 8,
               h3("Descriptive Statistics"),
               verbatimTextOutput("descResult")
             )
           )
  ),
  
  tabPanel("Usage",
           fluidPage(
             div(
               style = "text-align: center; padding: 50px;",
               h3("App Usage Statistics"),
               div(
                 style = "font-size: 24px; margin: 20px; padding: 20px; background-color: #f8f9fa; border-radius: 10px;",
                 paste("This app has been used", counter, "times.")
               ),
               p("Thank you for using CalcuStats!")
             )
           )
  )
)

server <- function(input, output, session) {  
  output$greetingTextCustom <- renderText({
    paste(getGreeting(), "Welcome to CalcuStats.")
  })
  
  # Custom Proportional Allocation section variables
  rv_custom <- reactiveValues(stratumCount = 1)
  
  observeEvent(input$addStratum_custom, { rv_custom$stratumCount <- rv_custom$stratumCount + 1 })
  observeEvent(input$removeStratum_custom, { if (rv_custom$stratumCount > 1) rv_custom$stratumCount <- rv_custom$stratumCount - 1 })
  
  output$stratumInputs_custom <- renderUI({
    lapply(1:rv_custom$stratumCount, function(i) {
      tagList(
        div(style = "margin-bottom: 10px; width: 100%;",
            textInput(paste0("stratum_custom", i),
                      label = paste("Stratum", i, "Name"),
                      value = paste("Stratum", i),
                      width = "100%")
        ),
        div(style = "margin-bottom: 20px; width = 100%;",
            numericInput(paste0("pop_custom", i),
                         label = paste("Stratum", i, "Population"),
                         value = 100,
                         min = 0,
                         width = "100%")
        ),
        tags$hr()
      )
    })
  })
  
  totalPopulation_custom <- reactive({
    sum(sapply(1:rv_custom$stratumCount, function(i) input[[paste0("pop_custom", i)]]), na.rm = TRUE)
  })
  
  adjustedSampleSize_custom <- reactive({
    n <- input$custom_sample
    if (n == 0) return(0)
    non_response_rate <- input$non_response_custom / 100
    ceiling(n / (1 - non_response_rate))
  })
  
  output$adjustedSampleSizeCustom <- renderText({
    n <- input$custom_sample
    adj_n <- adjustedSampleSize_custom()
    non_response_rate <- input$non_response_custom
    
    paste(
      "Adjusting for Non-Response:\n",
      "1. Original sample size: ", n, "\n",
      "2. Non-response rate: ", non_response_rate, "%\n",
      "3. Adjusted sample size formula: n_adjusted = n / (1 - non_response_rate)\n",
      "4. Calculation: ", n, " / (1 - ", non_response_rate/100, ") = ", n / (1 - non_response_rate/100), "\n",
      "5. Apply ceiling function: ", adj_n, "\n",
      "\nFinal adjusted sample size: ", adj_n
    )
  })
  
  allocationData_custom <- reactive({
    N <- totalPopulation_custom()
    n <- adjustedSampleSize_custom()
    strata <- sapply(1:rv_custom$stratumCount, function(i) input[[paste0("stratum_custom", i)]] )
    pops <- sapply(1:rv_custom$stratumCount, function(i) input[[paste0("pop_custom", i)]] )
    proportions <- pops / N
    raw_samples <- proportions * n
    rounded_samples <- floor(raw_samples)
    remainder <- n - sum(rounded_samples)
    
    if (remainder > 0) {
      decimal_parts <- raw_samples - rounded_samples
      indices <- order(decimal_parts, decreasing = TRUE)[1:remainder]
      rounded_samples[indices] <- rounded_samples[indices] + 1
    }
    
    calculations <- sapply(1:rv_custom$stratumCount, function(i) {
      paste0(
        "(", pops[i], " / ", N, ") × ", n, 
        " = ", round(proportions[i], 4), " × ", n, 
        " = ", round(raw_samples[i], 2), " → ", rounded_samples[i]
      )
    })
    
    df <- data.frame(
      Stratum = strata,
      Population = pops,
      Calculation = calculations,
      Proportional_Sample = rounded_samples
    )
    
    total_row <- data.frame(
      Stratum = "Total",
      Population = N,
      Calculation = paste0("Sum = ", sum(rounded_samples)),
      Proportional_Sample = sum(rounded_samples)
    )
    
    rbind(df, total_row)
  })
  
  interpretationText_custom <- reactive({
    alloc <- allocationData_custom()
    if (is.null(alloc)) return("")
    sentences <- paste0(alloc$Stratum[-nrow(alloc)],
                        " (population: ", alloc$Population[-nrow(alloc)],
                        ") should contribute ", alloc$Proportional_Sample[-nrow(alloc)],
                        " participants.")
    paste0("Based on your specified sample size of ", input$custom_sample,
           " and a non-response rate of ", input$non_response_custom, "%",
           ", the adjusted sample size is ", adjustedSampleSize_custom(), ". ",
           paste(sentences, collapse = " "))
  })
  
  output$totalPopCustom <- renderText({ paste("Total Population:", totalPopulation_custom()) })
  output$sampleSizeCustom <- renderText({ paste("Your Sample Size:", input$custom_sample) })
  output$allocationTableCustom <- renderTable({ allocationData_custom() })
  output$interpretationTextCustom <- renderText({ interpretationText_custom() })
  
  output$downloadCustomWord <- downloadHandler(
    filename = function() paste("Proportional_Allocation_Report_", Sys.Date(), ".docx", sep = ""),
    content = function(file) {
      doc <- read_docx() %>%
        body_add_par("CalcuStats", style = "heading 1") %>%
        body_add_par("Proportional Allocation for Specified Sample Size", style = "heading 2") %>%
        body_add_par(paste("Total Population:", totalPopulation_custom()), style = "Normal") %>%
        body_add_par(paste("Your Specified Sample Size:", input$custom_sample), style = "Normal") %>%
        body_add_par(paste("Non-response Rate:", input$non_response_custom, "%"), style = "Normal") %>%
        body_add_par(paste("Adjusted Sample Size (accounting for non-response):", adjustedSampleSize_custom()), style = "Normal") %>%
        body_add_par("Proportional Allocation Table:", style = "heading 2") %>%
        body_add_par("The proportional allocation formula is: n_i = (N_i / N) × n", style = "Normal")
      ft <- flextable(allocationData_custom()) %>% autofit()
      doc <- doc %>% body_add_flextable(ft)
      doc <- doc %>%
        body_add_par("Interpretation", style = "heading 2") %>%
        body_add_par(interpretationText_custom(), style = "Normal")
      print(doc, target = file)
    }
  )
  
  output$downloadCustomSteps <- downloadHandler(
    filename = function() paste("Proportional_Allocation_Steps_", Sys.Date(), ".txt", sep = ""),
    content = function(file) {
      steps <- c(
        "Proportional Allocation Calculation Steps",
        "=======================================",
        "",
        paste("Total Population (N):", totalPopulation_custom()),
        paste("Your Specified Sample Size (n):", input$custom_sample),
        paste("Non-response Rate:", input$non_response_custom, "%"),
        "",
        "1. Non-response Adjustment:",
        paste("Adjusted sample size = ", input$custom_sample, " / (1 - ", input$non_response_custom/100, ")"),
        paste("Calculation: ", input$custom_sample, " / ", (1 - input$non_response_custom/100), " = ", input$custom_sample / (1 - input$non_response_custom/100)),
        paste("Final adjusted sample size: ", adjustedSampleSize_custom()),
        "",
        "2. Proportional Allocation:",
        capture.output(print(allocationData_custom())),
        "",
        "Interpretation:",
        interpretationText_custom()
      )
      
      writeLines(steps, file)
    }
  )
  
  # Taro Yamane section variables
  rv <- reactiveValues(stratumCount = 1)
  
  observeEvent(input$addStratum, { rv$stratumCount <- rv$stratumCount + 1 })
  observeEvent(input$removeStratum, { if (rv$stratumCount > 1) rv$stratumCount <- rv$stratumCount - 1 })
  
  output$stratumInputs <- renderUI({
    lapply(1:rv$stratumCount, function(i) {
      tagList(
        div(style = "margin-bottom: 10px; width: 100%;",
            textInput(paste0("stratum", i),
                      label = paste("Stratum", i, "Name"),
                      value = paste("Stratum", i),
                      width = "100%")
        ),
        div(style = "margin-bottom: 20px; width = 100%;",
            numericInput(paste0("pop", i),
                         label = paste("Stratum", i, "Population"),
                         value = 100,
                         min = 0,
                         width = "100%")
        ),
        tags$hr()
      )
    })
  })
  
  totalPopulation <- reactive({
    sum(sapply(1:rv$stratumCount, function(i) input[[paste0("pop", i)]]), na.rm = TRUE)
  })
  
  sampleSize <- reactive({
    N <- totalPopulation()
    e <- input$e
    if (N == 0 || e == 0) return(0)
    ceiling(N / (1 + N * e^2))
  })
  
  adjustedSampleSize <- reactive({
    n <- sampleSize()
    if (n == 0) return(0)
    non_response_rate <- input$non_response / 100
    ceiling(n / (1 - non_response_rate))
  })
  
  output$formulaExplanation <- renderText({
    N <- totalPopulation()
    e <- input$e
    if (N == 0 || e == 0) return("")
    
    denominator <- 1 + N * e^2
    raw_sample <- N / denominator
    
    paste(
      "Yamane Formula Calculation Steps:\n",
      "1. Formula: n = N / (1 + N*e²)\n",
      "2. Total Population (N) = ", N, "\n",
      "3. Margin of Error (e) = ", e, "\n",
      "4. Calculate denominator: 1 + N*e² = 1 + (", N, " * ", e, "²) = 1 + (", N, " * ", e^2, ") = ", denominator, "\n",
      "5. Calculate raw sample size: ", N, " / ", denominator, " = ", raw_sample, "\n",
      "6. Apply ceiling function to round up to nearest integer: ", ceiling(raw_sample), "\n",
      "\nFinal calculated sample size after rounding: ", sampleSize()
    )
  })
  
  output$adjustedSampleSize <- renderText({
    n <- sampleSize()
    adj_n <- adjustedSampleSize()
    non_response_rate <- input$non_response
    
    paste(
      "Adjusting for Non-Response:\n",
      "1. Original sample size: ", n, "\n",
      "2. Non-response rate: ", non_response_rate, "%\n",
      "3. Adjusted sample size formula: n_adjusted = n / (1 - non_response_rate)\n",
      "4. Calculation: ", n, " / (1 - ", non_response_rate/100, ") = ", n / (1 - non_response_rate/100), "\n",
      "5. Apply ceiling function: ", adj_n, "\n",
      "\nFinal adjusted sample size: ", adj_n
    )
  })
  
  allocationData <- reactive({
    N <- totalPopulation()
    n <- adjustedSampleSize()
    strata <- sapply(1:rv$stratumCount, function(i) input[[paste0("stratum", i)]] )
    pops <- sapply(1:rv$stratumCount, function(i) input[[paste0("pop", i)]] )
    proportions <- pops / N
    raw_samples <- proportions * n
    rounded_samples <- floor(raw_samples)
    remainder <- n - sum(rounded_samples)
    
    if (remainder > 0) {
      decimal_parts <- raw_samples - rounded_samples
      indices <- order(decimal_parts, decreasing = TRUE)[1:remainder]
      rounded_samples[indices] <- rounded_samples[indices] + 1
    }
    
    calculations <- sapply(1:rv$stratumCount, function(i) {
      paste0(
        "(", pops[i], " / ", N, ") × ", n, 
        " = ", round(proportions[i], 4), " × ", n, 
        " = ", round(raw_samples[i], 2), " → ", rounded_samples[i]
      )
    })
    
    df <- data.frame(
      Stratum = strata,
      Population = pops,
      Calculation = calculations,
      Proportional_Sample = rounded_samples
    )
    
    total_row <- data.frame(
      Stratum = "Total",
      Population = N,
      Calculation = paste0("Sum = ", sum(rounded_samples)),
      Proportional_Sample = sum(rounded_samples)
    )
    
    rbind(df, total_row)
  })
  
  interpretationText <- reactive({
    alloc <- allocationData()
    if (is.null(alloc)) return("")
    sentences <- paste0(alloc$Stratum[-nrow(alloc)],
                        " (population: ", alloc$Population[-nrow(alloc)],
                        ") should contribute ", alloc$Proportional_Sample[-nrow(alloc)],
                        " participants.")
    paste0("Based on a total population of ", totalPopulation(),
           ", a margin of error of ", input$e,
           ", and a non-response rate of ", input$non_response, "%",
           ", the required sample size is ", adjustedSampleSize(), ". ",
           paste(sentences, collapse = " "))
  })
  
  output$totalPop <- renderText({ paste("Total Population:", totalPopulation()) })
  output$sampleSize <- renderText({ paste("Initial Sample Size:", sampleSize()) })
  output$allocationTable <- renderTable({ allocationData() })
  output$interpretationText <- renderText({ interpretationText() })
  
  output$downloadWord <- downloadHandler(
    filename = function() paste("Sample_Size_Report_", Sys.Date(), ".docx", sep = ""),
    content = function(file) {
      doc <- read_docx() %>%
        body_add_par("CalcuStats", style = "heading 1") %>%
        body_add_par("Sample Size Calculation (Taro Yamane Method)", style = "heading 2") %>%
        body_add_par(paste("Total Population:", totalPopulation()), style = "Normal") %>%
        body_add_par(paste("Margin of Error:", input$e), style = "Normal") %>%
        body_add_par(paste("Non-response Rate:", input$non_response, "%"), style = "Normal") %>%
        body_add_par(paste("Initial Sample Size:", sampleSize()), style = "Normal") %>%
        body_add_par(paste("Adjusted Sample Size (accounting for non-response):", adjustedSampleSize()), style = "Normal") %>%
        body_add_par("Proportional Allocation Table:", style = "heading 2") %>%
        body_add_par("The proportional allocation formula is: n_i = (N_i / N) × n", style = "Normal")
      ft <- flextable(allocationData()) %>% autofit()
      doc <- doc %>% body_add_flextable(ft)
      doc <- doc %>%
        body_add_par("Interpretation", style = "heading 2") %>%
        body_add_par(interpretationText(), style = "Normal")
      print(doc, target = file)
    }
  )
  
  output$downloadSteps <- downloadHandler(
    filename = function() paste("Taro_Yamane_Calculation_Steps_", Sys.Date(), ".txt", sep = ""),
    content = function(file) {
      steps <- c(
        "Taro Yamane Sample Size Calculation Steps",
        "========================================",
        "",
        paste("Total Population (N):", totalPopulation()),
        paste("Margin of Error (e):", input$e),
        paste("Non-response Rate:", input$non_response, "%"),
        "",
        "1. Basic Yamane Formula Calculation:",
        paste("Formula: n = N / (1 + N*e²)"),
        paste("Calculation: ", totalPopulation(), " / (1 + ", totalPopulation(), "*", input$e, "²)"),
        paste("Denominator: 1 + ", totalPopulation(), "*", input$e^2, " = ", 1 + totalPopulation() * input$e^2),
        paste("Raw sample size: ", totalPopulation(), " / ", (1 + totalPopulation() * input$e^2), " = ", totalPopulation() / (1 + totalPopulation() * input$e^2)),
        paste("Rounded sample size: ", sampleSize()),
        "",
        "2. Non-response Adjustment:",
        paste("Adjusted sample size = ", sampleSize(), " / (1 - ", input$non_response/100, ")"),
        paste("Calculation: ", sampleSize(), " / ", (1 - input$non_response/100), " = ", sampleSize() / (1 - input$non_response/100)),
        paste("Final adjusted sample size: ", adjustedSampleSize()),
        "",
        "3. Proportional Allocation:",
        capture.output(print(allocationData()))
      )
      
      writeLines(steps, file)
    }
  )
  
  # Cochran section variables
  rv_c <- reactiveValues(stratumCount = 1)
  
  observeEvent(input$addStratum_c, { rv_c$stratumCount <- rv_c$stratumCount + 1 })
  observeEvent(input$removeStratum_c, { if (rv_c$stratumCount > 1) rv_c$stratumCount <- rv_c$stratumCount - 1 })
  
  output$stratumInputs_c <- renderUI({
    lapply(1:rv_c$stratumCount, function(i) {
      tagList(
        div(style = "margin-bottom: 10px; width: 100%;",
            textInput(paste0("stratum_c", i),
                      label = paste("Stratum", i, "Name"),
                      value = paste("Stratum", i),
                      width = "100%")
        ),
        div(style = "margin-bottom: 20px; width = 100%;",
            numericInput(paste0("pop_c", i),
                         label = paste("Stratum", i, "Population"),
                         value = 100,
                         min = 0,
                         width = "100%")
        ),
        tags$hr()
      )
    })
  })
  
  totalPopulation_c <- reactive({
    sum(sapply(1:rv_c$stratumCount, function(i) input[[paste0("pop_c", i)]]), na.rm = TRUE)
  })
  
  cochranSampleSize <- reactive({
    p <- input$p
    z <- input$z
    e <- input$e_c
    N <- if (is.null(input$N_c) || is.na(input$N_c)) NULL else input$N_c
    
    n0 <- (z^2 * p * (1 - p)) / (e^2)
    
    if (!is.null(N) && N > 0) {
      n <- n0 / (1 + (n0 - 1)/N)
      return(list(initial = n0, corrected = n, finite_correction_applied = TRUE))
    } else {
      return(list(initial = n0, corrected = n0, finite_correction_applied = FALSE))
    }
  })
  
  adjustedCochranSampleSize <- reactive({
    res <- cochranSampleSize()
    n <- res$corrected
    if (n == 0) return(0)
    non_response_rate <- input$non_response_c / 100
    ceiling(n / (1 - non_response_rate))
  })
  
  cochranAllocationData <- reactive({
    N <- if (is.null(input$N_c) || is.na(input$N_c)) totalPopulation_c() else input$N_c
    n <- adjustedCochranSampleSize()
    strata <- sapply(1:rv_c$stratumCount, function(i) input[[paste0("stratum_c", i)]] )
    pops <- sapply(1:rv_c$stratumCount, function(i) input[[paste0("pop_c", i)]] )
    
    if (length(pops) == 0 || sum(pops) == 0) return(NULL)
    
    proportions <- pops / N
    raw_samples <- proportions * n
    rounded_samples <- floor(raw_samples)
    remainder <- n - sum(rounded_samples)
    
    if (remainder > 0) {
      decimal_parts <- raw_samples - rounded_samples
      indices <- order(decimal_parts, decreasing = TRUE)[1:remainder]
      rounded_samples[indices] <- rounded_samples[indices] + 1
    }
    
    calculations <- sapply(1:rv_c$stratumCount, function(i) {
      paste0(
        "(", pops[i], " / ", N, ") × ", n, 
        " = ", round(proportions[i], 4), " × ", n, 
        " = ", round(raw_samples[i], 2), " → ", rounded_samples[i]
      )
    })
    
    df <- data.frame(
      Stratum = strata,
      Population = pops,
      Calculation = calculations,
      Proportional_Sample = rounded_samples
    )
    
    total_row <- data.frame(
      Stratum = "Total",
      Population = N,
      Calculation = paste0("Sum = ", sum(rounded_samples)),
      Proportional_Sample = sum(rounded_samples)
    )
    
    rbind(df, total_row)
  })
  
  cochranInterpretationText <- reactive({
    res <- cochranSampleSize()
    p <- input$p
    z <- input$z
    e <- input$e_c
    N <- if (is.null(input$N_c) || is.na(input$N_c)) totalPopulation_c() else input$N_c
    alloc <- cochranAllocationData()
    
    if (is.null(alloc)) {
      interpretation <- paste(
        "Assuming a population proportion of", p,
        ", a confidence level corresponding to Z =", z,
        " (", round(pnorm(z) * 200 - 100, 1), "% confidence),",
        " and a margin of error of", e,
        ", the required sample size is", ceiling(res$corrected), "."
      )
      
      if (res$finite_correction_applied) {
        interpretation <- paste0(interpretation,
                                 " This includes a finite population correction for N = ", N, "."
        )
      } else {
        interpretation <- paste0(interpretation,
                                 " No finite population correction was applied (assumes infinite population)."
        )
      }
      
      return(interpretation)
    }
    
    sentences <- paste0(alloc$Stratum[-nrow(alloc)],
                        " (population: ", alloc$Population[-nrow(alloc)],
                        ") should contribute ", alloc$Proportional_Sample[-nrow(alloc)],
                        " participants.")
    
    interpretation <- paste(
      "Assuming a population proportion of", p,
      ", a confidence level corresponding to Z =", z,
      " (", round(pnorm(z) * 200 - 100, 1), "% confidence),",
      " and a margin of error of", e,
      ", the required sample size is", adjustedCochranSampleSize(), "."
    )
    
    if (res$finite_correction_applied) {
      interpretation <- paste0(interpretation,
                               " This includes a finite population correction for N = ", N, "."
      )
    } else {
      interpretation <- paste0(interpretation,
                               " No finite population correction was applied (assumes infinite population)."
      )
    }
    
    paste0(interpretation, " ", paste(sentences, collapse = " "))
  })
  
  output$cochranFormulaExplanation <- renderText({
    p <- input$p
    z <- input$z
    e <- input$e_c
    N <- if (is.null(input$N_c) || is.na(input$N_c)) NULL else input$N_c
    res <- cochranSampleSize()
    
    calculation_steps <- paste(
      "Cochran's Formula Calculation Steps:\n",
      "1. Basic formula: n₀ = (Z² × p × (1-p)) / e²\n",
      "2. Z-score (Z) = ", z, "\n",
      "3. Estimated proportion (p) = ", p, "\n",
      "4. Margin of error (e) = ", e, "\n",
      "5. Calculate numerator (Z² × p × (1-p)): (", z, "² × ", p, " × ", (1-p), ") = ", 
      (z^2 * p * (1 - p)), "\n",
      "6. Calculate denominator (e²): ", e, "² = ", e^2, "\n",
      "7. Initial sample size (n₀): ", (z^2 * p * (1 - p)), " / ", e^2, " = ", res$initial, "\n"
    )
    
    if (res$finite_correction_applied) {
      calculation_steps <- paste0(calculation_steps,
                                  "\nFinite Population Correction Applied (N = ", N, "):\n",
                                  "8. Correction formula: n = n₀ / (1 + (n₀ - 1)/N)\n",
                                  "9. Calculate denominator: 1 + (", res$initial, " - 1)/", N, " = ", 
                                  (1 + (res$initial - 1)/N), "\n",
                                  "10. Corrected sample size: ", res$initial, " / ", (1 + (res$initial - 1)/N), " = ", res$corrected, "\n"
      )
    } else {
      calculation_steps <- paste0(calculation_steps,
                                  "\nNo finite population correction applied (population size not specified or infinite).\n"
      )
    }
    
    calculation_steps
  })
  
  output$cochranAdjustedSample <- renderText({
    res <- cochranSampleSize()
    adj_n <- adjustedCochranSampleSize()
    non_response_rate <- input$non_response_c
    
    paste(
      "Adjusting for Non-Response:\n",
      "1. Original sample size: ", ceiling(res$corrected), "\n",
      "2. Non-response rate: ", non_response_rate, "%\n",
      "3. Adjusted sample size formula: n_adjusted = n / (1 - non_response_rate)\n",
      "4. Calculation: ", ceiling(res$corrected), " / (1 - ", non_response_rate/100, ") = ", ceiling(res$corrected) / (1 - non_response_rate/100), "\n",
      "5. Apply ceiling function: ", adj_n, "\n",
      "\nFinal adjusted sample size: ", adj_n
    )
  })
  
  output$cochranSample <- renderText({
    res <- cochranSampleSize()
    if (res$finite_correction_applied) {
      paste("Initial Sample Size (n₀):", ceiling(res$initial), "\n",
            "Corrected Sample Size (n):", ceiling(res$corrected))
    } else {
      paste("Sample Size (n₀):", ceiling(res$initial))
    }
  })
  
  output$cochranAllocationTable <- renderTable({
    cochranAllocationData()
  })
  
  output$cochranInterpretation <- renderText({
    cochranInterpretationText()
  })
  
  output$downloadCochranWord <- downloadHandler(
    filename = function() paste("Cochran_Sample_Size_Report_", Sys.Date(), ".docx", sep = ""),
    content = function(file) {
      res <- cochranSampleSize()
      p <- input$p
      z <- input$z
      e <- input$e_c
      N <- if (is.null(input$N_c) || is.na(input$N_c)) totalPopulation_c() else input$N_c
      alloc <- cochranAllocationData()
      
      doc <- read_docx() %>%
        body_add_par("CalcuStats", style = "heading 1") %>%
        body_add_par("Sample Size Calculation (Cochran's Method)", style = "heading 2") %>%
        body_add_par(paste("Estimated Proportion (p):", p), style = "Normal") %>%
        body_add_par(paste("Z-score (Z):", z), style = "Normal") %>%
        body_add_par(paste("Margin of Error (e):", e), style = "Normal") %>%
        body_add_par(paste("Non-response Rate:", input$non_response_c, "%"), style = "Normal")
      
      if (!is.null(N)) {
        doc <- doc %>% body_add_par(paste("Population Size (N):", N), style = "Normal")
      }
      
      doc <- doc %>%
        body_add_par(paste("Initial Sample Size (n₀):", ceiling(res$initial)), style = "Normal")
      
      if (res$finite_correction_applied) {
        doc <- doc %>%
          body_add_par(paste("Corrected Sample Size (n):", ceiling(res$corrected)), style = "Normal")
      }
      
      doc <- doc %>%
        body_add_par(paste("Adjusted Sample Size (accounting for non-response):", adjustedCochranSampleSize()), style = "Normal")
      
      if (!is.null(alloc)) {
        doc <- doc %>%
          body_add_par("Proportional Allocation Table:", style = "heading 2") %>%
          body_add_par("The proportional allocation formula is: n_i = (N_i / N) × n", style = "Normal")
        ft <- flextable(alloc) %>% autofit()
        doc <- doc %>% body_add_flextable(ft)
      }
      
      doc <- doc %>%
        body_add_par("Interpretation", style = "heading 2") %>%
        body_add_par(cochranInterpretationText(), style = "Normal")
      
      print(doc, target = file)
    }
  )
  
  output$downloadCochranSteps <- downloadHandler(
    filename = function() {
      paste("Cochran_Calculation_Steps_", Sys.Date(), ".txt", sep = "")
    },
    content = function(file) {
      res <- cochranSampleSize()
      p <- input$p
      z <- input$z
      e <- input$e_c
      N <- if (is.null(input$N_c) || is.na(input$N_c)) NULL else input$N_c
      alloc <- cochranAllocationData()
      
      steps <- c(
        "Cochran Sample Size Calculation Steps",
        "====================================",
        "",
        paste("Estimated Proportion (p):", p),
        paste("Z-score (Z):", z),
        paste("Margin of Error (e):", e),
        paste("Non-response Rate:", input$non_response_c, "%"),
        if (!is.null(N)) paste("Population Size (N):", N) else "Population Size: Infinite",
        "",
        "1. Basic Cochran Formula Calculation:",
        paste("Formula: n₀ = (Z² × p × (1-p)) / e²"),
        paste("Calculation: (", z, "² × ", p, " × ", (1-p), ") / ", e, "²"),
        paste("Numerator: ", z^2, " × ", p, " × ", (1-p), " = ", (z^2 * p * (1 - p))),
        paste("Denominator: ", e, "² = ", e^2),
        paste("Initial sample size (n₀): ", (z^2 * p * (1 - p)), " / ", e^2, " = ", res$initial)
      )
      
      if (res$finite_correction_applied) {
        steps <- c(steps,
                   "",
                   "2. Finite Population Correction:",
                   paste("Formula: n = n₀ / (1 + (n₀ - 1)/N)"),
                   paste("Calculation: ", res$initial, " / (1 + (", res$initial, " - 1)/", N, ")"),
                   paste("Denominator: 1 + (", res$initial, " - 1)/", N, " = ", (1 + (res$initial - 1)/N)),
                   paste("Corrected sample size: ", res$initial, " / ", (1 + (res$initial - 1)/N), " = ", res$corrected),
                   "",
                   paste("Final sample size before non-response adjustment: ", ceiling(res$corrected))
        )
      } else {
        steps <- c(steps,
                   "",
                   "2. No finite population correction applied (infinite population assumed)",
                   "",
                   paste("Final sample size before non-response adjustment: ", ceiling(res$initial))
        )
      }
      
      steps <- c(steps,
                 "",
                 "3. Non-response Adjustment:",
                 paste("Non-response rate: ", input$non_response_c, "%"),
                 paste("Adjusted sample size = ", ceiling(res$corrected), " / (1 - ", input$non_response_c/100, ")"),
                 paste("Calculation: ", ceiling(res$corrected), " / ", (1 - input$non_response_c/100), " = ", ceiling(res$corrected) / (1 - input$non_response_c/100)),
                 paste("Final adjusted sample size: ", adjustedCochranSampleSize())
      )
      
      if (!is.null(alloc)) {
        steps <- c(steps,
                   "",
                   "4. Proportional Allocation:",
                   capture.output(print(alloc))
        )
      }
      
      interpretation <- cochranInterpretationText()
      
      steps <- c(steps, "", "Interpretation:", interpretation)
      
      writeLines(steps, file)
    }
  )
  
  # Power Analysis section
  powerResult <- reactiveVal()
  
  observeEvent(input$runPower, {
    result <- switch(input$testType,
                     "Independent t-test" = pwr.t.test(d = input$effectSize, sig.level = input$alpha, power = input$power, type = "two.sample"),
                     "Paired t-test" = pwr.t.test(d = input$effectSize, sig.level = input$alpha, power = input$power, type = "paired"),
                     "One-sample t-test" = pwr.t.test(d = input$effectSize, sig.level = input$alpha, power = input$power, type = "one.sample"),
                     "One-Way ANOVA" = pwr.anova.test(f = input$effectSize, sig.level = input$alpha, power = input$power, k = 1),
                     "Two-Way ANOVA" = pwr.anova.test(f = input$effectSize, sig.level = input$alpha, power = input$power, k = 2),
                     "Proportion" = pwr.p.test(h = input$effectSize, sig.level = input$alpha, power = input$power),
                     "Correlation" = pwr.r.test(r = input$effectSize, sig.level = input$alpha, power = input$power),
                     "Chi-squared" = pwr.chisq.test(w = input$effectSize, sig.level = input$alpha, power = input$power, df = 1),
                     "Simple Linear Regression" = pwr.f2.test(u = 1, f2 = input$effectSize, sig.level = input$alpha, power = input$power),
                     "Multiple Linear Regression" = pwr.f2.test(u = input$predictors, f2 = input$effectSize, sig.level = input$alpha, power = input$power)
    )
    powerResult(result)
    output$powerResult <- renderPrint({ result })
  })
  
  output$downloadPowerSteps <- downloadHandler(
    filename = function() paste("Power_Analysis_Steps_", Sys.Date(), ".txt", sep = ""),
    content = function(file) {
      req(powerResult())
      result <- powerResult()
      
      test_info <- switch(input$testType,
                          "Independent t-test" = "Independent samples t-test",
                          "Paired t-test" = "Paired samples t-test",
                          "One-sample t-test" = "One-sample t-test",
                          "One-Way ANOVA" = "One-Way ANOVA",
                          "Two-Way ANOVA" = "Two-Way ANOVA",
                          "Proportion" = "Test of proportion",
                          "Correlation" = "Correlation test",
                          "Chi-squared" = "Chi-squared test",
                          "Simple Linear Regression" = "Simple Linear Regression",
                          "Multiple Linear Regression" = "Multiple Linear Regression"
      )
      
      formula_text <- switch(input$testType,
                             "Independent t-test" = "pwr.t.test(d = effect_size, sig.level = alpha, power = power, type = 'two.sample')",
                             "Paired t-test" = "pwr.t.test(d = effect_size, sig.level = alpha, power = power, type = 'paired')",
                             "One-sample t-test" = "pwr.t.test(d = effect_size, sig.level = alpha, power = power, type = 'one.sample')",
                             "One-Way ANOVA" = "pwr.anova.test(f = effect_size, sig.level = alpha, power = power, k = 1)",
                             "Two-Way ANOVA" = "pwr.anova.test(f = effect_size, sig.level = alpha, power = power, k = 2)",
                             "Proportion" = "pwr.p.test(h = effect_size, sig.level = alpha, power = power)",
                             "Correlation" = "pwr.r.test(r = effect_size, sig.level = alpha, power = power)",
                             "Chi-squared" = "pwr.chisq.test(w = effect_size, sig.level = alpha, power = power, df = 1)",
                             "Simple Linear Regression" = "pwr.f2.test(u = 1, f2 = effect_size, sig.level = alpha, power = power)",
                             "Multiple Linear Regression" = paste0("pwr.f2.test(u = ", input$predictors, ", f2 = effect_size, sig.level = alpha, power = power)")
      )
      
      steps <- c(
        "Power Analysis Calculation Steps",
        "===============================",
        "",
        paste("Test Type:", test_info),
        paste("Effect Size:", input$effectSize),
        paste("Significance Level (alpha):", input$alpha),
        paste("Desired Power:", input$power),
        if (input$testType == "Multiple Linear Regression") paste("Number of Predictors:", input$predictors) else "",
        "",
        "R Function Used:",
        formula_text,
        "",
        "Results:",
        capture.output(print(result)),
        "",
        "Interpretation:",
        paste("For a", test_info, "with an effect size of", input$effectSize, ","),
        paste("a significance level of", input$alpha, ","),
        paste("and desired power of", input$power, ","),
        paste("the required sample size is", result$n, ".")
      )
      
      writeLines(steps, file)
    }
  )
  
  # Descriptive Statistics section
  descResult <- reactiveVal()
  
  observeEvent(input$runDesc, {
    req(input$dataInput)
    lines <- unlist(strsplit(input$dataInput, "[\n\r]+"))
    
    if (length(lines) < 2) {
      output$descResult <- renderPrint({ "Please include header (Variable name) and at least one row of data." })
      return()
    }
    
    header <- lines[1]
    data_lines <- lines[-1]
    
    suppressWarnings({
      nums <- as.numeric(data_lines)
      is_numeric <- !is.na(nums)
    })
    
    if (all(is_numeric)) {
      values <- nums[is_numeric]
      n_miss <- sum(is.na(nums))
      desc <- list(
        Mean = mean(values),
        SD = sd(values),
        Min = min(values),
        Max = max(values),
        Mode = as.numeric(names(sort(table(values), decreasing = TRUE)[1])),
        Median = median(values),
        Skewness = skewness(values),
        Kurtosis = kurtosis(values),
        Missing = n_miss,
        Shapiro_Wilk_p = shapiro.test(values)$p.value
      )
      desc$total <- sum(values)
      
      descResult(list(type = "numeric", header = header, results = desc))
      
      output$descResult <- renderPrint({
        cat(paste("Descriptive Statistics for:", header, "\n"))
        cat("(Numeric Data)\n\n")
        print(desc)
      })
    } else {
      categories <- trimws(data_lines)
      tbl <- table(factor(categories))
      cum_freq <- cumsum(tbl)
      percent <- prop.table(tbl) * 100
      n_miss <- sum(is.na(categories))
      
      df <- data.frame(
        Category = names(tbl),
        Frequency = as.vector(tbl),
        Percentage = round(as.vector(percent), 2),
        Cumulative_Frequency = as.vector(cum_freq)
      )
      
      df <- rbind(df, data.frame(
        Category = "Total", 
        Frequency = sum(tbl), 
        Percentage = 100, 
        Cumulative_Frequency = max(cum_freq)
      ))
      
      descResult(list(type = "categorical", header = header, results = df, missing = n_miss))
      
      output$descResult <- renderPrint({
        cat(paste("Descriptive Statistics for:", header, "\n"))
        cat("(Categorical Data)\n\n")
        print(df)
        cat("\nMissing Values:", n_miss, "\n")
      })
    }
  })
  
  output$downloadDescSteps <- downloadHandler(
    filename = function() paste("Descriptive_Stats_Steps_", Sys.Date(), ".txt", sep = ""),
    content = function(file) {
      req(descResult())
      result <- descResult()
      
      if (result$type == "numeric") {
        steps <- c(
          "Descriptive Statistics Calculation Steps (Numeric Data)",
          "======================================================",
          "",
          paste("Variable Name:", result$header),
          "",
          "Calculated Statistics:",
          paste("Mean:", result$results$Mean),
          paste("Standard Deviation:", result$results$SD),
          paste("Minimum:", result$results$Min),
          paste("Maximum:", result$results$Max),
          paste("Median:", result$results$Median),
          paste("Mode:", result$results$Mode),
          paste("Skewness:", result$results$Skewness),
          paste("Kurtosis:", result$results$Kurtosis),
          paste("Total:", result$results$total),
          paste("Missing Values:", result$results$Missing),
          paste("Shapiro-Wilk p-value:", result$results$Shapiro_Wilk_p),
          "",
          "Interpretation:",
          ifelse(result$results$Shapiro_Wilk_p > 0.05, 
                 "The data appears to be normally distributed", 
                 "The data does not appear to be normally distributed"),
          paste("(Shapiro-Wilk test p-value:", result$results$Shapiro_Wilk_p, ")")
        )
      } else {
        steps <- c(
          "Descriptive Statistics Calculation Steps (Categorical Data)",
          "==========================================================",
          "",
          paste("Variable Name:", result$header),
          "",
          "Frequency Table:",
          capture.output(print(result$results)),
          "",
          paste("Missing Values:", result$missing)
        )
      }
      
      writeLines(steps, file)
    }
  )
}

shinyApp(ui = ui, server = server)
