library(ggplot2)
library(dplyr)
library(openxlsx)

wb <- openxlsx::createWorkbook()

# Insert Tabs
addWorksheet(wb, "Summary", gridLines = FALSE)
addWorksheet(wb, "Issues", gridLines = FALSE)

issues <- open_issues %>% 
  select(Task.Name, Description, Assigned.To) %>% 
  rename("Issue" = Task.Name,
         "Assigned To" = Assigned.To)

# Populate with data
openxlsx::writeData(wb = wb, x = wk_report, startCol = 4, startRow = 4, borders = "rows", sheet = "Summary")
openxlsx::writeData(wb = wb, x = issues, startCol = 1, startRow = 4, borders = "rows", sheet = "Issues")

# Format layouts and styles

setColWidths(wb, "Summary", cols = c(4, 5, 6, 7, 8, 9, 10, 11, 12), widths = c(10, 13, 13, 13, 13, 13, 13, 13, 13))

setColWidths(wb, "Issues", cols = c(1, 2, 3), widths = c(30, 55, 18))

## create and add a style to the column headers
headerStyle <- openxlsx::createStyle(
  fontSize = 11, fontColour = "#FFFFFF", halign = "left",
  fgFill = ppa_cols[2], border = "TopBottom", borderColour = "#4F81BD", wrapText = TRUE
)

## style for body
bodyStyle <- openxlsx::createStyle(
  border = "TopBottom", borderColour = "#4F81BD", valign = "center", halign = "left", wrapText = TRUE
)

highlight.row <- openxlsx::createStyle(
  fontSize = 11, fontColour = "#FFFFFF", halign = "right",
  fgFill = ppa_cols[1], border = "TopBottom", borderColour = "#4F81BD", wrapText = TRUE
)

underline.header <- openxlsx::createStyle( # Create a green bar at the top of the page
  fontSize = 11, fontColour = "#FFFFFF", halign = "right",
  fgFill = "#FFFFFF", border = "Top", borderColour = ppa_cols[2]
)

openxlsx::addStyle(wb, sheet = "Summary", style = bodyStyle, rows = 4, cols = c(4:11), gridExpand = TRUE)
openxlsx::addStyle(wb, sheet = "Summary", style = headerStyle, rows = 4, cols = c(4:11), gridExpand = TRUE)
openxlsx::addStyle(wb, sheet = "Summary", style = highlight.row, rows = 6, cols = c(4:11), gridExpand = TRUE)
openxlsx::addStyle(wb, sheet = "Summary", style = underline.header, rows = 1, cols = c(1:15), gridExpand = TRUE)

openxlsx::addStyle(wb, sheet = "Issues", style = bodyStyle, rows = c(5:(4 + nrow(issues))), cols = c(1:3), gridExpand = TRUE)
openxlsx::addStyle(wb, sheet = "Issues", style = underline.header, rows = 1, cols = c(1:3), gridExpand = TRUE)
openxlsx::addStyle(wb, sheet = "Issues", style = headerStyle, rows = 4, cols = c(1:3), gridExpand = TRUE)


print(tasks_over_time)
insertPlot(wb, sheet = 1, xy = c("C", 11), width = 15, height = 10, fileType = "png", units = "cm")

print(late_versus_ontime)
insertPlot(wb, sheet = 1, xy = c("I", 11), width = 15, height = 10, fileType = "png", units = "cm")

print(plot_tasks_assigned_to)
insertPlot(wb, sheet = 1, xy = c("F", 31), width = 15, height = 9, fileType = "png", units = "cm")

# Format Summary page layout
openxlsx::pageSetup(wb, sheet = "Summary", orientation = "landscape", paperSize = 8)
openxlsx::pageSetup(wb, sheet = "Issues", orientation = "portrait", fitToWidth = TRUE)

# setHeader(wb, "TASK REPORT", position = "left")
setHeaderFooter(wb, sheet = "Summary", 
                header = c('&"Arial"&B&14&K008C98TASK REPORT', 
                                                  NA, 
                                                  NA),
                footer = c("Printed On: &[Date]", NA, "Page &[Page] of &[Pages]"),)

setHeaderFooter(wb, sheet = "Issues", 
                header = c('&"Arial"&B&14&K008C98ISSUES REGISTER', 
                           NA, 
                           NA),
                footer = c("Printed On: &[Date]", NA, "Page &[Page] of &[Pages]"),)


saveWorkbook(wb, here::here("output", paste0(Sys.Date(), " Task Report.xlsx")), overwrite = TRUE, returnValue = FALSE)
