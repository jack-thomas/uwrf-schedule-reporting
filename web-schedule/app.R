#
# UWRF Web Reporting
#

library(shiny)
library(rvest)
library(magrittr)
library(stringr)
library(openxlsx)

# PRE-PROCESSING
## SETUP TERM INPUT
terms <- read_html("https://www.uwrf.edu/ClassSchedule/") %>%
  html_nodes(., xpath = '//*[(@id = "selectTerm")]') %>%
  html_children() %>%
  html_attr(., "value")
term_names <- read_html("https://www.uwrf.edu/ClassSchedule/") %>%
  html_nodes(., xpath = '//*[(@id = "selectTerm")]') %>%
  html_text(.) %>%
  gsub("\\\r", "", .) %>%
  gsub("\\\n", "", .) %>%
  gsub("\\\t+", ",", .) %>%
  strsplit(., ",")
term_names <- term_names[[1]]
names(terms) <- term_names

## SETUP SUBJECT INPUT
subjects <- read_html("https://www.uwrf.edu/ClassSchedule/") %>%
  html_nodes(., xpath = '//*[(@id = "selectSubject")]') %>%
  html_children() %>%
  html_attr(., "value")
subject_names <- read_html("https://www.uwrf.edu/ClassSchedule/") %>%
  html_nodes(., xpath = '//*[(@id = "selectSubject")]') %>%
  html_text(.) %>%
  gsub(",", "", .) %>%
  gsub("\\\r\\\n\\\t\\\t\\\t\\\t\\\t\\\t\\\t\\\t\\\t\\\t\\\t", ",", .) %>%
  gsub("\\\r\\\n\\\t\\\t\\\t\\\t\\\t\\\t\\\t\\\t\\\t\\\t", "", .) %>%
  strsplit(., ",")
subject_names <- subject_names[[1]]
names(subjects) <- subject_names

## GET WEB RESULTS
toRender <- function(term, subject, result = "master"){
  # The result input is backwards... ``result = "courses"`` returns a summary.
  session1 <- html_session("https://www.uwrf.edu/ClassSchedule/")
  form1 <- html_form(read_html(session1))[[3]]
  results <- submit_form(session1, set_values(form1, courseTerm = term, courseSubject = subject, instructor = "")) %>%
    read_html() %>%
    html_nodes(., xpath = '//*[(@id = "courseTable")]') %>%
    as.character() %>%
    gsub("[\\\r\\\n\\\t]*", '', .) %>%
    gsub('\\\"', '', .) %>%
    gsub("\\s+", " ", .) %>%
    str_extract_all(., "<td(.*?)</td>") %>%
    .[[1]] %>%
    gsub("<(.*?)>", "", .) %>%
    str_trim()
  master <- data.frame(matrix(0, nrow = length(which(results == "Section")), ncol = 9))
  names(master) <- c("CatalogNumber", "Title", "Credits", "Section", "ClassNumber", "Instructor", "Enrolled", "Room", "Time")
  master$Section <- which(results == "Section")
  master$CatalogNumber <- (which(results == "Catalog Number"))[findInterval(master$Section, which(results == "Catalog Number"))]
  master$Instructor <- (which(results == "Instructor"))[findInterval(master$Section, which(results == "Instructor")) + 1]
  master$Enrolled <- (which(results == "Enrolled"))[findInterval(master$Section, which(results == "Enrolled")) + 1]
  room1 <- results[(which(results == "Room")) + 1]
  room2 <- findInterval(master$Section, which(results == "Room")) + 1
  room3 <- data.frame(room2[findInterval(c(1:length(room1)), room2)], room1)
  names(room3) <- c("x1", "x2")
  room4 <- aggregate(room3, by = list(room3$x1), FUN = toString)[,3] %>%
    str_split(., ",") %>%
    gsub("c\\(", "", .) %>%
    gsub("\"", "", .) %>%
    gsub(")", "", .) %>%
    gsub("  ", "", .) %>%
    str_split(., ",")
  compare <- function(v) all(sapply( as.list(v[-1]), FUN=function(z) {identical(z, v[1])}))
  first <- function(a) {a[1]}
  jPaste <- function(a) {toString(a) %>% gsub(",", ";", .)}
  master$Room <- ifelse(sapply(room4, compare), sapply(room4, first), sapply(room4, jPaste))
  master$ClassNumber <- (which(results == "Class Number"))[findInterval(master$Section, which(results == "Class Number")) + 1]
  time1 <- results[(which(results == "Time")) + 1]
  time2 <- findInterval(master$Section, which(results == "Time")) + 1
  time3 <- data.frame(time2[findInterval(c(1:length(time1)), time2)], time1)
  names(time3) <- c("x1", "x2")
  time4 <- aggregate(time3, by = list(time3$x1), FUN = toString)[,3]
  master$Time <- gsub(",", ";", time4)
  master$Section <- results[master$Section + 1]
  master$Title <- results[master$CatalogNumber + 6]
  master$Title <- gsub("&amp;", "&", master$Title)
  master$Credits <- results[master$CatalogNumber + 7]
  master$CatalogNumber <- results[master$CatalogNumber + 4]
  master$Instructor <- results[master$Instructor + 1]
  master$Instructor <- gsub(",", ", ", master$Instructor)
  master$Enrolled <- results[master$Enrolled + 1]
  master$ClassNumber <- results[master$ClassNumber + 1]
  names(master) <- c("Catalog Number", "Course Title", "Credits", "Section Number",
                     "Class Number", "Instructor", "Enrollment", "Room(s)", "Days and Times")
  # Get summary data before moving around columns...
  courses <- master[, c(1, 7)]
  # Requested order: Cat Num, Sec, Title, Cred, Times, Enroll, Room, Instr, Cl Num
  master <- master[, c(1, 4, 2, 3, 9, 7, 8, 6, 5)]
  # Continue computing summary
  courses[, 1] <- as.numeric(courses[, 1])
  courses[, 2] <- as.numeric(str_extract(courses[, 2], "[0-9]*"))
  courses[, 3] <- rep(1, nrow(courses))
  names(courses) <- c("c", "e", "n")
  courses <- aggregate(cbind(courses$n, courses$e), by = list(Category = courses$c), FUN = sum)
  names(courses) <- c("Catalog Number", "Sections", "Enrollment")
  
  if (result == "courses") return(courses)
  else return(master)
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
          downloadButton("esisDownload")
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
    filename = function() {return(paste("eSIS ", input$esisTerm, ".xlsx", sep = ''))},
    content = function(file) {xlsx::write.xlsx(esisTable(input$pasted),
                                               sheetName = "Courses",
                                               file, row.names = FALSE)}
  )
  
  # WEB
  output$webTable <- renderTable(toRender(input$TermInput, input$SubjInput), digits = 0)
  output$webTableSummary <- renderTable(toRender(input$TermInput, input$SubjInput, result = "courses"), digits = 0)
  output$webDownload <- downloadHandler(
    filename = function() {
      return(
        paste(ifelse(input$TermInput == "", "All", term_names[which(terms == input$TermInput)]),
              ' Web.xlsx', sep = '')
      )
    },
    content = function(file) {
      xlsx::write.xlsx(toRender(input$TermInput, input$SubjInput),
                       sheetName = "Courses",
                       file, row.names = FALSE)
      xlsx::write.xlsx(toRender(input$TermInput, input$SubjInput, result = "courses"),
                       sheetName = "Summary",
                       file, row.names = FALSE, append = TRUE)
    }
  )
}

shinyApp(uwrf_ui, uwrf_server)
