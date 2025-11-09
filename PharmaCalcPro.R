# app.R

library(shiny)
library(shinythemes)
library(dplyr)
library(validate)
library(rmarkdown)
library(pagedown)

# Define UI with professional layout
ui <- fluidPage(
  theme = shinytheme("lumen"),
  tags$head(
    tags$style(HTML("
      body {
        font-size: 16px;
        color: #333333;
        background-color: #f8f9fa;
      }
      .navbar {
        background-color: #005b96 !important;
        border: none;
        font-size: 18px;
      }
      .navbar-brand {
        font-weight: bold;
        font-size: 22px;
        color: white !important;
      }
      .nav-tabs>li>a {
        color: #ffffff !important;
        background-color: #005b96;
        font-weight: bold;
        font-size: 17px;
        border: none !important;
        margin-right: 2px;
        border-radius: 5px 5px 0 0;
        padding: 12px 20px;
      }
      .nav-tabs>li.active>a {
        color: white !important;
        background-color: #0077b6 !important;
        border: none !important;
      }
      .nav-tabs>li>a:hover {
        color: white !important;
        background-color: #0096c7 !important;
      }
      .footer {
        background-color: #005b96;
        color: white;
        padding: 20px;
        margin-top: 30px;
        text-align: center;
        font-size: 14px;
      }
      .well {
        box-shadow: 0 2px 10px rgba(0,0,0,0.05);
        background-color: white;
        border-radius: 8px;
        border: 1px solid #e0e0e0;
      }
      .result-box {
        background-color: #f1f9ff;
        border-left: 5px solid #0077b6;
        padding: 20px;
        margin-bottom: 25px;
        border-radius: 5px;
        font-size: 17px;
        color: #333333;
      }
      .instruction-box {
        background-color: #f0f8f5;
        border-left: 5px solid #28a745;
        padding: 20px;
        margin-bottom: 25px;
        border-radius: 5px;
        font-size: 17px;
        color: #333333;
      }
      .conversion-box {
        background-color: #fff9e6;
        border-left: 5px solid #ffc107;
        padding: 20px;
        margin-bottom: 25px;
        border-radius: 5px;
        font-size: 17px;
        color: #333333;
      }
      h1, h2, h3, h4 {
        color: #005b96;
        font-weight: 600;
      }
      .btn-primary {
        background-color: #0077b6;
        border-color: #005b96;
        font-size: 16px;
        padding: 10px 16px;
        font-weight: 500;
      }
      .btn-default {
        font-size: 16px;
        padding: 10px 16px;
        font-weight: 500;
      }
      .btn-success {
        background-color: #28a745;
        border-color: #1e7e34;
        font-size: 16px;
        padding: 10px 16px;
        font-weight: 500;
      }
      .form-control {
        font-size: 16px;
        height: 42px;
        border: 1px solid #ced4da;
      }
      .selectize-input {
        font-size: 16px;
        height: 42px;
        border: 1px solid #ced4da;
      }
      table {
        font-size: 16px;
      }
      .nav-tabs {
        border-bottom: 3px solid #0077b6;
      }
      .shiny-input-container {
        color: #495057;
      }
      label {
        font-weight: 500;
        color: #495057;
      }
      .concentration-inputs {
        display: flex;
        gap: 10px;
        align-items: end;
      }
      .concentration-inputs .form-group {
        flex: 1;
        margin-bottom: 0;
      }
    "))
  ),
  
  navbarPage(
    title = div(img(src = "https://cdn-icons-png.flaticon.com/512/2008/2008740.png", 
                    height = "30px", 
                    style = "margin-right:10px;"),
                "PharmaCalc Pro"),
    id = "main_navbar",
    
    tabPanel(
      "Standard Calculator",
      sidebarLayout(
        sidebarPanel(
          width = 4,
          div(class = "well",
              h4("Drug Information", icon("flask", style = "color: #0077b6;")),
              div(class = "concentration-inputs",
                  div(
                    numericInput("drug_amount", "Drug Amount", 
                                 value = 1, min = 0.0001, step = 0.1),
                    selectInput("amount_unit", "Amount Unit", 
                                choices = c("mg", "mcg", "IU", "U", "mmol", "g", "mEq", "ng", "μg"),
                                selected = "mg")
                  ),
                  div(
                    numericInput("drug_volume", "Per Volume (mL)", 
                                 value = 1, min = 0.0001, step = 0.1),
                    p("mL", style = "font-weight: bold; margin-bottom: 15px;")
                  )
              ),
              
              hr(style = "border-top: 1px solid #e0e0e0;"),
              h4("Dosage Calculation", icon("calculator", style = "color: #0077b6;")),
              radioButtons("calculation_type", "Calculation Type",
                           choices = c("Calculate volume to administer" = "vol_to_admin",
                                       "Calculate dose from volume" = "dose_from_vol"),
                           selected = "vol_to_admin"),
              
              conditionalPanel(
                condition = "input.calculation_type == 'vol_to_admin'",
                numericInput("desired_dose", "Desired Dose", value = 1, min = 0.0001, step = 0.1),
                selectInput("dose_unit", "Dose Unit", 
                            choices = c("mg", "mcg", "IU", "U", "g", "mmol", 
                                        "mEq", "ng", "μg", "kg", "lb", "Custom"),
                            selected = "mg"),
                conditionalPanel(
                  condition = "input.dose_unit == 'Custom'",
                  textInput("custom_dose_unit", "Enter Custom Dose Unit", value = "")
                )
              ),
              
              conditionalPanel(
                condition = "input.calculation_type == 'dose_from_vol'",
                numericInput("volume_to_admin", "Volume to Administer (mL)", value = 1, min = 0.0001, step = 0.1)
              ),
              
              actionButton("calculate", "Calculate", 
                           icon = icon("play-circle"),
                           class = "btn-primary btn-block",
                           style = "margin-top: 15px;"),
              actionButton("reset", "Reset", 
                           icon = icon("redo"),
                           class = "btn-default btn-block",
                           style = "margin-top: 10px;"),
              conditionalPanel(
                condition = "input.calculate > 0",
                downloadButton("", "",
                               class = "btn-success btn-block",
                               style = "margin-top: 10px;")
              )
          )
        ),
        
        mainPanel(
          width = 8,
          conditionalPanel(
            condition = "input.calculate > 0",
            div(class = "result-box",
                h4("Calculation Result", icon("prescription-bottle-alt", style = "color: #0077b6;")),
                htmlOutput("result_text"),
                tableOutput("calculation_details")
            )
          ),
          div(class = "instruction-box",
              h4("Quick Guide", icon("info-circle", style = "color: #28a745;")),
              p("Enter the drug amount and volume to calculate concentration. Then choose whether to calculate volume to administer or dose from volume."),
              p(strong("Examples:")),
              tags$ul(
                tags$li("For Furosemide 20mg/2mL and desired dose of 40 mg:"),
                tags$ul(
                  tags$li("Enter Drug Amount: 20, Amount Unit: mg"),
                  tags$li("Enter Per Volume: 2 mL"),
                  tags$li("Desired Dose: 40 mg → Volume to administer: 4 mL")
                ),
                tags$li("For 2.5 mL of a 10,000 IU/5mL solution:"),
                tags$ul(
                  tags$li("Enter Drug Amount: 10000, Amount Unit: IU"),
                  tags$li("Enter Per Volume: 5 mL"),
                  tags$li("Volume to administer: 2.5 mL → Total dose: 5,000 IU")
                )
              )
          )
        )
      )
    ),
    
    tabPanel(
      "Weight-Based Dosing",
      sidebarLayout(
        sidebarPanel(
          width = 4,
          div(class = "well",
              h4("Patient Information", icon("user", style = "color: #0077b6;")),
              numericInput("patient_weight", "Patient Weight", 
                           value = 70, min = 0.1, step = 0.1),
              selectInput("weight_unit", "Weight Unit",
                          choices = c("kg", "lb"),
                          selected = "kg"),
              
              hr(style = "border-top: 1px solid #e0e0e0;"),
              h4("Dosage Information", icon("syringe", style = "color: #0077b6;")),
              numericInput("dose_per_kg", "Dosage (per kg)", 
                           value = 1, min = 0.0001, step = 0.1),
              selectInput("dose_unit_wt", "Dosage Unit",
                          choices = c("mg/kg", "mcg/kg", "IU/kg", "U/kg"),
                          selected = "mg/kg"),
              
              selectInput("frequency", "Frequency of Dose",
                          choices = c("Once daily" = "q24hr",
                                      "Every 12 hours (BID)" = "q12hr",
                                      "Every 8 hours (TID)" = "q8hr",
                                      "Every 6 hours (QID)" = "q6hr",
                                      "Every 4 hours" = "q4hr",
                                      "Every 3 hours" = "q3hr",
                                      "Every 2 hours" = "q2hr",
                                      "Every 1 hour" = "q1hr",
                                      "Other frequency" = "other"),
                          selected = "q12hr"),
              
              hr(style = "border-top: 1px solid #e0e0e0;"),
              h4("Liquid Formulation (Optional)", icon("flask", style = "color: #0077b6;")),
              div(class = "concentration-inputs",
                  div(
                    numericInput("liquid_med_amount", "Medication amount", 
                                 value = 0, min = 0.0001, step = 0.1),
                    selectInput("liquid_amount_unit", "Amount Unit",
                                choices = c("mg", "mcg", "IU", "U", "g"),
                                selected = "mg")
                  ),
                  div(
                    numericInput("liquid_med_volume", "Per volume (mL)", 
                                 value = 0, min = 0.0001, step = 0.1),
                    p("mL", style = "font-weight: bold; margin-bottom: 15px;")
                  )
              ),
              
              actionButton("calculate_wt", "Calculate", 
                           icon = icon("play-circle"),
                           class = "btn-primary btn-block",
                           style = "margin-top: 15px;"),
              actionButton("reset_wt", "Reset", 
                           icon = icon("redo"),
                           class = "btn-default btn-block",
                           style = "margin-top: 10px;"),
              conditionalPanel(
                condition = "input.calculate_wt > 0",
                downloadButton("download_weight", "Download PDF Report",
                               class = "btn-success btn-block",
                               style = "margin-top: 10px;")
              )
          )
        ),
        
        mainPanel(
          width = 8,
          conditionalPanel(
            condition = "input.calculate_wt > 0",
            div(class = "result-box",
                h4("Weight-Based Dosing Result", icon("prescription-bottle-alt", style = "color: #0077b6;")),
                htmlOutput("result_text_wt"),
                tableOutput("calculation_details_wt")
            )
          ),
          div(class = "instruction-box",
              h4("Quick Guide", icon("info-circle", style = "color: #28a745;")),
              p("Enter patient weight and dosage per kg to calculate the appropriate dose."),
              p(strong("Examples:")),
              tags$ul(
                tags$li("For a 34.5 kg patient with 20 mg/kg dosage (BID):"),
                tags$ul(
                  tags$li("Total daily dose: 1,380 mg (690 mg per dose)"),
                  tags$li("If using 20mg/12mL solution: 414 mL per dose")
                ),
                tags$li("For a 150 lb patient with 5 mcg/kg dosage (TID):"),
                tags$ul(
                  tags$li("Convert weight to kg (68.04 kg)"),
                  tags$li("Total daily dose: 1,021 mcg (340 mcg per dose)")
                )
              )
          )
        )
      )
    ),
    
    tabPanel(
      "Tablet Dosage",
      sidebarLayout(
        sidebarPanel(
          width = 4,
          div(class = "well",
              h4("Tablet Information", icon("pills", style = "color: #0077b6;")),
              numericInput("tablet_strength", "Tablet Strength", 
                           value = 100, min = 0.0001, step = 0.1),
              selectInput("tablet_unit", "Tablet Unit", 
                          choices = c("mg", "mcg", "IU", "U", "g", "mmol", 
                                      "mEq", "ng", "μg", "Custom"),
                          selected = "mg"),
              conditionalPanel(
                condition = "input.tablet_unit == 'Custom'",
                textInput("custom_tablet_unit", "Enter Custom Tablet Unit", value = "")
              ),
              
              hr(style = "border-top: 1px solid #e0e0e0;"),
              h4("Prescribed Dosage", icon("prescription", style = "color: #0077b6;")),
              numericInput("prescribed_dose", "Prescribed Dose", 
                           value = 200, min = 0.0001, step = 0.1),
              selectInput("prescribed_unit", "Dose Unit", 
                          choices = c("mg", "mcg", "IU", "U", "g", "mmol", 
                                      "mEq", "ng", "μg", "Custom"),
                          selected = "mg"),
              conditionalPanel(
                condition = "input.prescribed_unit == 'Custom'",
                textInput("custom_prescribed_unit", "Enter Custom Dose Unit", value = "")
              ),
              
              hr(style = "border-top: 1px solid #e0e0e0;"),
              checkboxInput("allow_partial", "Allow partial tablets?", value = TRUE),
              
              actionButton("calculate_tab", "Calculate", 
                           icon = icon("play-circle"),
                           class = "btn-primary btn-block",
                           style = "margin-top: 15px;"),
              actionButton("reset_tab", "Reset", 
                           icon = icon("redo"),
                           class = "btn-default btn-block",
                           style = "margin-top: 10px;"),
              conditionalPanel(
                condition = "input.calculate_tab > 0",
                downloadButton("download_tablet", "Download PDF Report",
                               class = "btn-success btn-block",
                               style = "margin-top: 10px;")
              )
          )
        ),
        
        mainPanel(
          width = 8,
          conditionalPanel(
            condition = "input.calculate_tab > 0",
            div(class = "result-box",
                h4("Tablet Dosage Result", icon("prescription-bottle-alt", style = "color: #0077b6;")),
                htmlOutput("result_text_tab"),
                tableOutput("calculation_details_tab")
            )
          ),
          div(class = "instruction-box",
              h4("Quick Guide", icon("info-circle", style = "color: #28a745;")),
              p("Enter the tablet strength and prescribed dose to calculate how many tablets to administer."),
              p(strong("Examples:")),
              tags$ul(
                tags$li("For 500 mg tablets and prescribed dose of 750 mg:"),
                tags$ul(
                  tags$li("With partial tablets: 1.5 tablets"),
                  tags$li("Without partial tablets: 2 tablets (1.5 rounded up)")
                ),
                tags$li("For 50 mcg tablets and prescribed dose of 125 mcg:"),
                tags$ul(
                  tags$li("With partial tablets: 2.5 tablets"),
                  tags$li("Without partial tablets: 3 tablets (2.5 rounded up)")
                )
              )
          )
        )
      )
    ),
    
    tabPanel(
      "Unit Reference",
      div(class = "conversion-box",
          h3("Unit Conversion Reference", icon("exchange-alt", style = "color: #ffc107;")),
          tableOutput("unit_conversion_table")
      ),
      div(class = "conversion-box",
          h3("Percentage Solutions", icon("percentage", style = "color: #ffc107;")),
          p("Percentage solutions are expressed as weight per volume (w/v):"),
          tableOutput("percentage_table")
      ),
      div(class = "conversion-box",
          h3("Common Medication Units", icon("pills", style = "color: #ffc107;")),
          tableOutput("medication_units_table")
      )
    ),
    
    tabPanel(
      "About",
      div(class = "instruction-box",
          h3("About PharmaCalc Pro", icon("question-circle", style = "color: #28a745;")),
          p("PharmaCalc Pro is a professional pharmacological dosage calculator designed for healthcare professionals, pharmacists, and medical students."),
          h4("Features:"),
          tags$ul(
            tags$li("Accurate dosage calculations for various units (mg, mcg, IU, mEq, etc.)"),
            tags$li("Weight-based dosing calculations"),
            tags$li("Tablet dosage calculations"),
            tags$li("Support for custom units"),
            tags$li("Two calculation modes: volume to administer or dose from volume"),
            tags$li("Comprehensive unit conversion reference"),
            tags$li("Support for over 5 different measurement units"),
            tags$li("PDF report generation for documentation")
          ),
          h4("Developer:"),
          p(strong("Mudasir Mohammed Ibrahim")),
          p("For suggestions or information, please contact:"),
          p(tags$a(href = "mailto:mudassiribrahim30@gmail.com", 
                   "mudassiribrahim30@gmail.com",
                   style = "color: #0077b6;")),
          p(tags$a(href = "https://github.com/mudassiribrahim30/", 
                   "GitHub Profile", 
                   target = "_blank",
                   style = "color: #0077b6;"))
      )
    )
  ),
  
  div(class = "footer",
      p("© 2025 PharmaCalc Pro | Developed by Mudasir Mohammed Ibrahim"),
      p("For educational and professional use only")
  )
)

# Server logic
server <- function(input, output, session) {
  
  # Reactive values for standard calculator
  rv <- reactiveValues(
    result = NULL,
    calculation_details = NULL,
    concentration = NULL
  )
  
  # Reactive values for weight-based calculator
  rv_wt <- reactiveValues(
    result = NULL,
    calculation_details = NULL
  )
  
  # Reactive values for tablet calculator
  rv_tab <- reactiveValues(
    result = NULL,
    calculation_details = NULL
  )
  
  # Reset standard calculator inputs
  observeEvent(input$reset, {
    updateNumericInput(session, "drug_amount", value = 1)
    updateNumericInput(session, "drug_volume", value = 1)
    updateSelectInput(session, "amount_unit", selected = "mg")
    updateRadioButtons(session, "calculation_type", selected = "vol_to_admin")
    updateNumericInput(session, "desired_dose", value = 1)
    updateSelectInput(session, "dose_unit", selected = "mg")
    updateNumericInput(session, "volume_to_admin", value = 1)
    updateTextInput(session, "custom_dose_unit", value = "")
    rv$result <- NULL
    rv$calculation_details <- NULL
    rv$concentration <- NULL
  })
  
  # Reset weight-based calculator inputs
  observeEvent(input$reset_wt, {
    updateNumericInput(session, "patient_weight", value = 70)
    updateSelectInput(session, "weight_unit", selected = "kg")
    updateNumericInput(session, "dose_per_kg", value = 1)
    updateSelectInput(session, "dose_unit_wt", selected = "mg/kg")
    updateSelectInput(session, "frequency", selected = "q12hr")
    updateNumericInput(session, "liquid_med_amount", value = 0)
    updateNumericInput(session, "liquid_med_volume", value = 0)
    rv_wt$result <- NULL
    rv_wt$calculation_details <- NULL
  })
  
  # Reset tablet calculator inputs
  observeEvent(input$reset_tab, {
    updateNumericInput(session, "tablet_strength", value = 100)
    updateSelectInput(session, "tablet_unit", selected = "mg")
    updateNumericInput(session, "prescribed_dose", value = 200)
    updateSelectInput(session, "prescribed_unit", selected = "mg")
    updateTextInput(session, "custom_tablet_unit", value = "")
    updateTextInput(session, "custom_prescribed_unit", value = "")
    updateCheckboxInput(session, "allow_partial", value = TRUE)
    rv_tab$result <- NULL
    rv_tab$calculation_details <- NULL
  })
  
  # Calculate for standard calculator
  observeEvent(input$calculate, {
    tryCatch({
      # Validate inputs
      req(input$drug_amount > 0, input$drug_volume > 0,
          ifelse(input$calculation_type == "vol_to_admin", 
                 input$desired_dose > 0, 
                 input$volume_to_admin > 0))
      
      # Calculate concentration
      concentration <- input$drug_amount / input$drug_volume
      rv$concentration <- concentration
      
      # Get units
      amount_unit <- input$amount_unit
      dose_unit <- ifelse(input$calculation_type == "vol_to_admin" && input$dose_unit == "Custom", 
                          input$custom_dose_unit, 
                          ifelse(input$calculation_type == "vol_to_admin", 
                                 input$dose_unit, 
                                 amount_unit))
      
      # Perform calculation
      if (input$calculation_type == "vol_to_admin") {
        # Volume to administer = Desired dose / Concentration
        volume <- input$desired_dose / concentration
        rv$result <- paste0(
          "Administer ", round(volume, 4), " mL",
          " of the ", input$drug_amount, " ", amount_unit, "/", input$drug_volume, " mL solution",
          " to achieve a dose of ", input$desired_dose, " ", dose_unit, "."
        )
        
        rv$calculation_details <- data.frame(
          Parameter = c("Drug Concentration", "Desired Dose", "Volume to Administer"),
          Value = c(paste(round(concentration, 4), amount_unit, "/mL"),
                    paste(input$desired_dose, dose_unit),
                    paste(round(volume, 4), "mL")),
          Formula = c(paste(input$drug_amount, "/", input$drug_volume), 
                      "", 
                      "Volume = Desired Dose / Concentration")
        )
      } else {
        # Dose = Volume * Concentration
        dose <- input$volume_to_admin * concentration
        rv$result <- paste0(
          input$volume_to_admin, " mL",
          " of the ", input$drug_amount, " ", amount_unit, "/", input$drug_volume, " mL solution",
          " contains ", round(dose, 4), " ", dose_unit, "."
        )
        
        rv$calculation_details <- data.frame(
          Parameter = c("Drug Concentration", "Volume to Administer", "Total Dose"),
          Value = c(paste(round(concentration, 4), amount_unit, "/mL"),
                    paste(input$volume_to_admin, "mL"),
                    paste(round(dose, 4), dose_unit)),
          Formula = c(paste(input$drug_amount, "/", input$drug_volume), 
                      "", 
                      "Dose = Volume × Concentration")
        )
      }
    }, error = function(e) {
      rv$result <- paste("Error:", e$message)
      rv$calculation_details <- NULL
    })
  })
  
  # Calculate for weight-based dosing
  observeEvent(input$calculate_wt, {
    tryCatch({
      # Validate inputs
      req(input$patient_weight > 0, input$dose_per_kg > 0)
      
      # Convert weight to kg if needed
      weight_kg <- ifelse(input$weight_unit == "kg", 
                          input$patient_weight, 
                          input$patient_weight * 0.453592)
      
      # Calculate total dose
      total_dose <- weight_kg * input$dose_per_kg
      
      # Calculate per dose based on frequency
      freq_info <- switch(input$frequency,
                          "q24hr" = list(doses_per_day = 1, label = "Once daily"),
                          "q12hr" = list(doses_per_day = 2, label = "Every 12 hours (BID)"),
                          "q8hr" = list(doses_per_day = 3, label = "Every 8 hours (TID)"),
                          "q6hr" = list(doses_per_day = 4, label = "Every 6 hours (QID)"),
                          "q4hr" = list(doses_per_day = 6, label = "Every 4 hours"),
                          "q3hr" = list(doses_per_day = 8, label = "Every 3 hours"),
                          "q2hr" = list(doses_per_day = 12, label = "Every 2 hours"),
                          "q1hr" = list(doses_per_day = 24, label = "Every 1 hour"),
                          list(doses_per_day = 1, label = "Per dose"))
      
      dose_per_admin <- total_dose / freq_info$doses_per_day
      
      # Calculate liquid volume if information provided
      liquid_info <- NULL
      if (!is.null(input$liquid_med_amount) && !is.null(input$liquid_med_volume) &&
          input$liquid_med_amount > 0 && input$liquid_med_volume > 0) {
        concentration <- input$liquid_med_amount / input$liquid_med_volume
        liquid_volume <- dose_per_admin / concentration
        liquid_info <- paste0(round(liquid_volume, 2), " mL")
      }
      
      # Prepare result
      rv_wt$result <- paste0(
        "For a ", input$patient_weight, " ", input$weight_unit, " patient (", round(weight_kg, 2), " kg):<br>",
        "<strong>Total daily dose:</strong> ", round(total_dose, 4), " ", sub("/kg", "", input$dose_unit_wt), "<br>",
        "<strong>Dose per administration (", freq_info$label, "):</strong> ", round(dose_per_admin, 4), " ", sub("/kg", "", input$dose_unit_wt),
        ifelse(!is.null(liquid_info), paste0("<br><strong>Liquid volume to administer:</strong> ", liquid_info), "")
      )
      
      # Prepare calculation details
      details <- data.frame(
        Parameter = c("Patient Weight", 
                      "Converted Weight (kg)", 
                      "Dosage",
                      "Total Daily Dose",
                      "Dose per Administration",
                      "Frequency"),
        Value = c(paste(input$patient_weight, input$weight_unit),
                  paste(round(weight_kg, 2), "kg"),
                  paste(input$dose_per_kg, input$dose_unit_wt),
                  paste(round(total_dose, 4), sub("/kg", "", input$dose_unit_wt)),
                  paste(round(dose_per_admin, 4), sub("/kg", "", input$dose_unit_wt)),
                  freq_info$label),
        Formula = c("", 
                    ifelse(input$weight_unit == "kg", "Same as input", "lb × 0.453592 = kg"),
                    "",
                    "Weight (kg) × Dosage",
                    "Total Daily Dose / Number of Doses",
                    "")
      )
      
      if (!is.null(liquid_info)) {
        details <- rbind(details,
                         data.frame(
                           Parameter = c("Liquid Concentration", "Volume to Administer"),
                           Value = c(paste(input$liquid_med_amount, input$liquid_amount_unit, "/", 
                                           input$liquid_med_volume, "mL"),
                                     liquid_info),
                           Formula = c(paste(input$liquid_med_amount, "/", input$liquid_med_volume),
                                       "Dose / Concentration")
                         ))
      }
      
      rv_wt$calculation_details <- details
      
    }, error = function(e) {
      rv_wt$result <- paste("Error:", e$message)
      rv_wt$calculation_details <- NULL
    })
  })
  
  # Calculate for tablet dosage
  observeEvent(input$calculate_tab, {
    tryCatch({
      # Validate inputs
      req(input$tablet_strength > 0, input$prescribed_dose > 0)
      
      # Get units
      tablet_unit <- ifelse(input$tablet_unit == "Custom", 
                            input$custom_tablet_unit, 
                            input$tablet_unit)
      
      prescribed_unit <- ifelse(input$prescribed_unit == "Custom", 
                                input$custom_prescribed_unit, 
                                input$prescribed_unit)
      
      # Check if units match (simplified - in real app would need conversion logic)
      if (tablet_unit != prescribed_unit) {
        stop("Tablet strength unit and prescribed dose unit must match")
      }
      
      # Calculate number of tablets
      num_tablets <- input$prescribed_dose / input$tablet_strength
      
      if (input$allow_partial) {
        result_text <- paste0(
          "Administer ", round(num_tablets, 2), " tablets ",
          "(prescribed dose: ", input$prescribed_dose, " ", prescribed_unit, 
          " / tablet strength: ", input$tablet_strength, " ", tablet_unit, ")."
        )
        rounded_text <- ""
      } else {
        rounded_num <- ceiling(num_tablets)
        result_text <- paste0(
          "Administer ", rounded_num, " whole tablets ",
          "(prescribed dose: ", input$prescribed_dose, " ", prescribed_unit, 
          " / tablet strength: ", input$tablet_strength, " ", tablet_unit, ")."
        )
        rounded_text <- paste0(
          "Note: Exact calculation was ", round(num_tablets, 2), " tablets, ",
          "rounded up to ", rounded_num, " whole tablets."
        )
      }
      
      rv_tab$result <- paste0(
        result_text,
        ifelse(nchar(rounded_text) > 0, paste0("<br>", rounded_text), "")
      )
      
      # Prepare calculation details
      details <- data.frame(
        Parameter = c("Tablet Strength", 
                      "Prescribed Dose", 
                      "Number of Tablets",
                      "Allow Partial Tablets?"),
        Value = c(paste(input$tablet_strength, tablet_unit),
                  paste(input$prescribed_dose, prescribed_unit),
                  ifelse(input$allow_partial, 
                         paste(round(num_tablets, 2), "tablets"),
                         paste(ceiling(num_tablets), "whole tablets (rounded up from", round(num_tablets, 2), ")")),
                  ifelse(input$allow_partial, "Yes", "No")),
        Formula = c("", 
                    "",
                    "Prescribed Dose / Tablet Strength",
                    "")
      )
      
      rv_tab$calculation_details <- details
      
    }, error = function(e) {
      rv_tab$result <- paste("Error:", e$message)
      rv_tab$calculation_details <- NULL
    })
  })
  
  # Render results for standard calculator
  output$result_text <- renderUI({
    if (!is.null(rv$result)) {
      HTML(paste0("<div style='font-size:17px;'>", rv$result, "</div>"))
    }
  })
  
  output$calculation_details <- renderTable({
    req(rv$calculation_details)
    rv$calculation_details
  }, align = 'l', bordered = TRUE, striped = TRUE, width = "100%")
  
  # Render results for weight-based calculator
  output$result_text_wt <- renderUI({
    if (!is.null(rv_wt$result)) {
      HTML(paste0("<div style='font-size:17px;'>", rv_wt$result, "</div>"))
    }
  })
  
  output$calculation_details_wt <- renderTable({
    req(rv_wt$calculation_details)
    rv_wt$calculation_details
  }, align = 'l', bordered = TRUE, striped = TRUE, width = "100%")
  
  # Render results for tablet calculator
  output$result_text_tab <- renderUI({
    if (!is.null(rv_tab$result)) {
      HTML(paste0("<div style='font-size:17px;'>", rv_tab$result, "</div>"))
    }
  })
  
  output$calculation_details_tab <- renderTable({
    req(rv_tab$calculation_details)
    rv_tab$calculation_details
  }, align = 'l', bordered = TRUE, striped = TRUE, width = "100%")
  
  # Download handlers for PDF reports
  output$download_standard <- downloadHandler(
    filename = function() {
      paste("standard-calculator-report-", Sys.Date(), ".pdf", sep = "")
    },
    content = function(file) {
      # Create a temporary Rmd file
      temp_rmd <- tempfile(fileext = ".Rmd")
      
      # Write the Rmd content
      writeLines(c(
        "---",
        "title: \"Standard Calculator Report\"",
        "output: pdf_document",
        "---",
        "",
        "## Calculation Result",
        "",
        paste(rv$result),
        "",
        "## Calculation Details",
        "",
        "```{r, echo=FALSE}",
        "rv$calculation_details",
        "```",
        "",
        "## Input Parameters",
        "",
        paste("- Drug Amount:", input$drug_amount, input$amount_unit),
        paste("- Drug Volume:", input$drug_volume, "mL"),
        paste("- Calculation Type:", ifelse(input$calculation_type == "vol_to_admin", 
                                            "Calculate volume to administer", 
                                            "Calculate dose from volume")),
        if(input$calculation_type == "vol_to_admin") {
          paste("- Desired Dose:", input$desired_dose, 
                ifelse(input$dose_unit == "Custom", input$custom_dose_unit, input$dose_unit))
        } else {
          paste("- Volume to Administer:", input$volume_to_admin, "mL")
        },
        "",
        "---",
        paste("Report generated on:", Sys.Date()),
        "PharmaCalc Pro - For professional use only"
      ), temp_rmd)
      
      # Render to PDF
      render(temp_rmd, output_file = file, quiet = TRUE)
    }
  )
  
  output$download_weight <- downloadHandler(
    filename = function() {
      paste("weight-based-report-", Sys.Date(), ".pdf", sep = "")
    },
    content = function(file) {
      # Create a temporary Rmd file
      temp_rmd <- tempfile(fileext = ".Rmd")
      
      # Write the Rmd content
      writeLines(c(
        "---",
        "title: \"Weight-Based Dosing Report\"",
        "output: pdf_document",
        "---",
        "",
        "## Calculation Result",
        "",
        paste(gsub("<br>", "  \n", gsub("<[^>]+>", "", rv_wt$result))),
        "",
        "## Calculation Details",
        "",
        "```{r, echo=FALSE}",
        "rv_wt$calculation_details",
        "```",
        "",
        "## Input Parameters",
        "",
        paste("- Patient Weight:", input$patient_weight, input$weight_unit),
        paste("- Dosage:", input$dose_per_kg, input$dose_unit_wt),
        paste("- Frequency:", input$frequency),
        if(input$liquid_med_amount > 0 && input$liquid_med_volume > 0) {
          paste("- Liquid Formulation:", input$liquid_med_amount, input$liquid_amount_unit, "/", 
                input$liquid_med_volume, "mL")
        },
        "",
        "---",
        paste("Report generated on:", Sys.Date()),
        "PharmaCalc Pro - For professional use only"
      ), temp_rmd)
      
      # Render to PDF
      render(temp_rmd, output_file = file, quiet = TRUE)
    }
  )
  
  output$download_tablet <- downloadHandler(
    filename = function() {
      paste("tablet-dosage-report-", Sys.Date(), ".pdf", sep = "")
    },
    content = function(file) {
      # Create a temporary Rmd file
      temp_rmd <- tempfile(fileext = ".Rmd")
      
      # Write the Rmd content
      writeLines(c(
        "---",
        "title: \"Tablet Dosage Report\"",
        "output: pdf_document",
        "---",
        "",
        "## Calculation Result",
        "",
        paste(gsub("<br>", "  \n", gsub("<[^>]+>", "", rv_tab$result))),
        "",
        "## Calculation Details",
        "",
        "```{r, echo=FALSE}",
        "rv_tab$calculation_details",
        "```",
        "",
        "## Input Parameters",
        "",
        paste("- Tablet Strength:", input$tablet_strength, 
              ifelse(input$tablet_unit == "Custom", input$custom_tablet_unit, input$tablet_unit)),
        paste("- Prescribed Dose:", input$prescribed_dose, 
              ifelse(input$prescribed_unit == "Custom", input$custom_prescribed_unit, input$prescribed_unit)),
        paste("- Allow Partial Tablets:", ifelse(input$allow_partial, "Yes", "No")),
        "",
        "---",
        paste("Report generated on:", Sys.Date()),
        "PharmaCalc Pro - For professional use only"
      ), temp_rmd)
      
      # Render to PDF
      render(temp_rmd, output_file = file, quiet = TRUE)
    }
  )
  
  # Unit conversion table
  output$unit_conversion_table <- renderTable({
    data.frame(
      "Conversion" = c("1 gram (g) = 1000 milligrams (mg)", 
                       "1 milligram (mg) = 1000 micrograms (mcg)", 
                       "1 liter (L) = 1000 milliliters (mL)",
                       "1 kilogram (kg) = 1000 grams (g)",
                       "1 pound (lb) = 0.453592 kilograms (kg)",
                       "1 International Unit (IU) Vitamin A ≈ 0.3 mcg retinol",
                       "1 IU Vitamin D = 0.025 mcg cholecalciferol",
                       "1 IU Vitamin E ≈ 0.67 mg d-alpha-tocopherol",
                       "1 IU Insulin = ~0.0347 mg human insulin",
                       "1 milliequivalent (mEq) ≈ 1 millimole (mmol) for monovalent ions",
                       "1 part per million (ppm) = 1 mg/L",
                       "1 part per billion (ppb) = 1 mcg/L"),
      "Notes" = c("", "", "", "", "",
                  "Vitamin A conversion varies by form",
                  "", 
                  "Vitamin E conversion varies by form",
                  "Exact conversion depends on insulin type",
                  "Adjust for valence of ion",
                  "For dilute aqueous solutions",
                  "For dilute aqueous solutions")
    )
  }, align = 'l', bordered = TRUE, striped = TRUE, width = "100%")
  
  # Percentage solution table
  output$percentage_table <- renderTable({
    data.frame(
      "Percentage" = c("1%", "2%", "0.5%", "5%", "10%", "20%", "50%"),
      "Concentration" = c("10 mg/mL", "20 mg/mL", "5 mg/mL", "50 mg/mL", "100 mg/mL", "200 mg/mL", "500 mg/mL"),
      "Example" = c("1% lidocaine = 10 mg/mL", 
                    "2% lidocaine = 20 mg/mL", 
                    "0.5% bupivacaine = 5 mg/mL",
                    "5% glucose = 50 mg/mL",
                    "10% calcium gluconate = 100 mg/mL",
                    "20% lipid emulsion = 200 mg/mL",
                    "50% dextrose = 500 mg/mL")
    )
  }, align = 'l', bordered = TRUE, striped = TRUE, width = "100%")
  
  # Medication units table
  output$medication_units_table <- renderTable({
    data.frame(
      "Unit" = c("mg (milligram)", "mcg (microgram)", "IU (International Unit)", 
                 "U (Unit)", "mEq (milliequivalent)", "mmol (millimole)",
                 "g (gram)", "ng (nanogram)", "μg (microgram)", "kg (kilogram)",
                 "lb (pound)", "ppm (parts per million)", "ppb (parts per billion)"),
      "Common Uses" = c("Most solid medications", "Highly potent drugs, hormones", 
                        "Vitamins, biologics, heparin", "Insulin, some enzymes", 
                        "Electrolytes (K+, Na+, etc.)", "Chemical substances",
                        "Bulk powders, some infusions", "Very potent substances",
                        "Alternative to mcg", "Patient weight", "Patient weight (US)",
                        "Trace elements in solutions", "Ultra-trace elements")
    )
  }, align = 'l', bordered = TRUE, striped = TRUE, width = "100%")
}

# Run the application

shinyApp(ui = ui, server = server)
