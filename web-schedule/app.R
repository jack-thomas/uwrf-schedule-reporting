#
# UWRF Web Reporting
#

library(shiny)
library(dplyr)
library(httr2)
library(jsonlite)
library(magrittr)
library(stringr)
library(openxlsx)

## GET WEB RESULTS
getData <- function(query, base_url = "https://students.uwrf.edu/custom/schedulelookup/") {
  resp <- request(base_url) %>%
    req_url_query(!!!query) %>%
    req_headers(Accept = "application/json", Referer = base_url) %>%
    req_perform() %>%
    resp_body_string() %>%
    fromJSON() %>%
    return()
}
## get background / supporting data
terms_data <- getData(query = list(f = "terms"))
terms <- terms_data$termId
names(terms) <- terms_data$termDesc
subjects_data <- getData(query = list(f = "subjects"))
subjects <- subjects_data$subjectId
names(subjects) <- subjects_data$subjectDesc
## get classes
toRender <- function(term, subject, result = "Courses") {
  # result can either be "Courses" (the default) or "Summary" (the second tab)
  result_df <- getData(query = list(courseTerm = term, courseSubject = subject))
  if (is.null(nrow(result_df))) {
    return(data.frame())
  }
  result_df %<>%
    mutate(
      time_mutated = ifelse(
        or(is.na(timeStart), is.na(timeEnd)),
        NA,
        gsub("\\s+", " ", str_trim(paste(
          paste(timeStart, timeEnd, sep = " - "),
          ifelse(and(!is.na(monday), monday == "Y"), "M", ""),
          ifelse(and(!is.na(tuesday), tuesday == "Y"), "Tu", ""),
          ifelse(and(!is.na(wednesday), wednesday == "Y"), "W", ""),
          ifelse(and(!is.na(thursday), thursday == "Y"), "Th", ""),
          ifelse(and(!is.na(friday), friday == "Y"), "F", ""),
          ifelse(and(!is.na(saturday), saturday == "Y"), "Sa", ""),
          ifelse(and(!is.na(sunday), sunday == "Y"), "Su", ""),
          sep = " "
        )))
      )
    ) %>%
    mutate(
      instructor_mutated = str_replace(instructor, ",", ", ")
    ) %>%
    select(
      `Catalog Number` = catalogNumber,
      `Course Title` = description,
      Credits = unitsMax,
      `Section Number` = section,
      `Class Number` = classNumber,
      Instructor = instructor_mutated,
      Enrollment = enrollTotal, #TODO consider adding x of y instead of just x
      `Room(s)` = location,
      `Time(s)` = time_mutated
    ) %>%
    group_by(`Catalog Number`, `Course Title`, Credits, `Section Number`, `Class Number`) %>%
    summarize(
      `Instructor(s)` = paste(unique(Instructor), collapse = "; "),
      Enrollment = mean(Enrollment),
      `Room(s)` = paste(unique(`Room(s)`), collapse = "; "),
      `Time(s)` = gsub("^$", NA, paste(unique(na.omit(`Time(s)`)), collapse = "; ")),
      .groups = "rowwise"
    ) %>%
    arrange(
      readr::parse_number(`Catalog Number`),
      readr::parse_number(`Section Number`)
    )
  if (result == "Summary") {
    result_df %<>%
      group_by(`Catalog Number`) %>%
      summarize(
        Sections = n_distinct(`Section Number`),
        Enrollment = sum(Enrollment)
      )
  }
  return(result_df)
}

## GET ESIS RESULTS
esisTable <- function(longText){
  allitems <- str_split(longText, "\n")
  allitems <- allitems[[1]]
  a <- data.frame(matrix(0, nrow = length(which(substr(allitems, 1, 5) == "Class")), ncol = 2))
  a[, 1] <- which(substr(allitems, 1, 5) == "Class")
  a[, 2] <- which(substr(allitems, 1, 8) %in% c("Open", "Closed"))
  a$Class <- allitems[a[, 1] + 1] %>%
    as.numeric()
  a$Number <- c(rep(0, nrow(a)))
  a$Name <- c(rep(0, nrow(a)))
  a$Section <- paste(allitems[a[, 1] + 2], allitems[a[, 1] + 3])
  a$DaysTimes <- ifelse(
    substr(allitems[a[, 1] + 5], 1, 2) %in% c("Mo", "Tu", "We", "Th", "Fr", "Sa", "Su"),
    ifelse(
      allitems[a[, 1] + 4] == allitems[a[, 1] + 5],
      allitems[a[, 1] + 4],
      paste0(allitems[a[, 1] + 4], "; ", allitems[a[, 1] + 5])
    ),
    allitems[a[, 1] + 4]
  )
  a$Rooms <- ifelse(
    substr(allitems[a[, 1] + 5], 1, 2) %in% c("Mo", "Tu", "We", "Th", "Fr", "Sa", "Su"),
    ifelse(
      allitems[a[, 1] + 6] == allitems[a[, 1] + 7],
      allitems[a[, 1] + 6],
      paste0(allitems[a[, 1] + 6], "; ", allitems[a[, 1] + 7])
    ),
    allitems[a[, 1] + 5]
  )
  a$Instructor <- ifelse(
    substr(allitems[a[, 1] + 5], 1, 2) %in% c("Mo", "Tu", "We", "Th", "Fr", "Sa", "Su"),
    ifelse(
      allitems[a[, 1] + 8] == allitems[a[, 1] + 9],
      allitems[a[, 1] + 8],
      paste0(allitems[a[, 1] + 8], "; ", allitems[a[, 1] + 9])
    ),
    allitems[a[, 1] + 6]
  )
  a$Name <- findInterval(a[, 1], which(substr(allitems, 1, 8) == "Collapse")) %>%
    which(substr(allitems, 1, 8) == "Collapse")[.] %>%
    allitems[.] %>%
    gsub("\\s+", " ", .) %>%
    gsub("Collapse section ", "", .)
  a$Number <- str_extract(a$Name, "MATH [0-9]* - ") %>%
    gsub("MATH ", "", .) %>%
    gsub(" - ", "", .) %>%
    as.numeric()
  a$Name <- gsub("MATH [0-9]* - ", "", a$Name) %>%
    substr(., 1, (nchar(.) - 1) / 2)
  a <- a[, c(4, 6, 5, 7, 8, 9, 3)]
  names(a) <- c("Catalog Number", "Section Number", "Course Title", "Days and Times",
                "Room(s)", "Instructor", "Class Number")
  return(a)
}

uwrf_ui <- fluidPage(
  titlePanel("UWRF Web Schedule Reporting"),
  sidebarLayout(
    sidebarPanel(
      selectInput("Source", "Source", list("Web" = "web", "eSIS" = "esis"), selected = "web"),
      conditionalPanel(
        condition = "input.Source == 'esis'",
        textInput("esisTerm", "Term"),
        conditionalPanel(
          condition = "input.esisTerm == ''",
          textOutput("esisNoTerm")
        ),
        conditionalPanel(
          condition = "input.esisTerm != ''",
          downloadButton(outputId = "esisDownload")
        )
      ),
      conditionalPanel(
        condition = "input.Source == 'web'",
        selectInput("TermInput", "Term", terms),
        selectInput("SubjInput", "Subject", subjects, selected = "MATH"),
        downloadButton(outputId = "webDownload")
      )
    ),
    mainPanel(
      conditionalPanel(
        condition = "input.Source == 'web'",
        tabsetPanel(
          tabPanel("Courses", tableOutput("webTable")),
          tabPanel("Summary", tableOutput("webTableSummary"))
        )
      ),
      conditionalPanel(
        condition = "input.Source == 'esis'",
        textAreaInput("pasted", "Paste Data Here", width = '100%'),
        tableOutput("esisTable")
      )
    )
  )
)

uwrf_server <- function(input, output) {
  #ESIS
  output$esisNoTerm <- renderText("Please provide a term to download.")
  output$esisTable <- renderTable(esisTable(input$pasted), digits = 0)
  output$esisDownload <- downloadHandler(
    filename = function() {return(paste(input$esisTerm, " eSIS Schedule Report.xlsx", sep = ''))},
    content = function(file) {
      xlsx::write.xlsx2(esisTable(input$pasted), sheetName = "Courses", file, row.names = FALSE)
    }
  )

  # WEB
  output$webTable <- renderTable(toRender(input$TermInput, input$SubjInput), digits = 0)
  output$webTableSummary <- renderTable(toRender(input$TermInput, input$SubjInput, "Summary"), digits = 0)
  output$webDownload <- downloadHandler(
    filename = function() {
      return(
        paste(
          ifelse(input$TermInput == "", "All", terms_data$termDesc[which(terms == input$TermInput)]),
          ' Web Schedule Report.xlsx',
          sep = ''
        )
      )
    },
    content = function(file) {
      xlsx::write.xlsx2(
        as.data.frame(toRender(input$TermInput, input$SubjInput, "Courses")),
        sheetName = "Courses", file, row.names = FALSE
      )
      xlsx::write.xlsx2(
        as.data.frame(toRender(input$TermInput, input$SubjInput, "Summary")),
        sheetName = "Summary", file, row.names = FALSE,
        append = TRUE
      )
    }
  )
}

shinyApp(uwrf_ui, uwrf_server)
