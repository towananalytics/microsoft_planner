library(ggplot2)
library(dplyr)
library(openxlsx)

wb <- openxlsx::createWorkbook()

specific.labels <- "ToPH" # Case sensitive multiple labels can be included by using | operator -> "ToPH|Issue|ExCo"
sheet.names <- c("Task Summary", "Issues", "ToPH Action Items", "All Tasks")
start.cols <- c(4, 1, 2, 2) # Where the tables should start for each of the above sheets
start.rows <- c(4, 4, 4, 4)
sheet1.name <- "Task Summary"
sheet2.name <- "Issues"

issues <- open_issues %>% 
  select(Task.Name, Description, Assigned.To) %>% 
  rename("Issue" = Task.Name,
         "Assigned To" = Assigned.To)

wk_report <- wk_report %>%
  select(-"Completion Percentage")

# Extract actions for specific labels (specific.labels)
specific_actions <- latest_plan %>% 
  filter(stringr::str_detect(Labels, specific.labels, negate = FALSE),
         Progress != "Completed") %>% 
  select(Task.Name, Description, Progress, Assigned.To, Due.Date) 
  

plan <- latest_plan %>% select(Task.Name, Description, Progress, Assigned.To, Due.Date)

report_list <- list(wk_report, issues, specific_actions, plan)

## create and add a style to the column headers
headerStyle <- openxlsx::createStyle(
  fontSize = 11, fontColour = "#FFFFFF", halign = "center", valign = "center",
  fgFill = ppa_cols[2], border = "TopBottom", borderColour = "#4F81BD", wrapText = TRUE
)

## style for body
bodyStyle <- openxlsx::createStyle(
  border = "TopBottom", borderColour = "#4F81BD", valign = "center", halign = "center", wrapText = TRUE
)

## create and add a style to the column headers
headerStyle.issues <- openxlsx::createStyle(
  fontSize = 11, fontColour = "#FFFFFF", halign = "left", valign = "center",
  fgFill = ppa_cols[2], border = "TopBottom", borderColour = "#4F81BD", wrapText = TRUE
)

## style for body
# bodyStyle.issues <- openxlsx::createStyle(
#   border = "TopBottom", borderColour = "#4F81BD", valign = "center", halign = "left", wrapText = TRUE
# )

highlight.row <- openxlsx::createStyle(
  fontSize = 11, fontColour = "#FFFFFF", halign = "center",
  fgFill = ppa_cols[1], border = "TopBottom", borderColour = "#4F81BD", wrapText = TRUE
)

underline.header <- openxlsx::createStyle( # Create a green bar at the top of the page
  fontSize = 11, fontColour = "#FFFFFF", halign = "center",
  fgFill = "#FFFFFF", border = "Top", borderColour = ppa_cols[2]
)

percentage.style <- openxlsx::createStyle(numFmt = "PERCENTAGE", valign = "center", halign = "center",
)

style.index.rows <- 1
style.index.cols <- 1

for(i in seq_along(sheet.names)){
  
  addWorksheet(wb, sheet.names[i], gridLines = FALSE)
  openxlsx::writeData(wb = wb, x = report_list[i][[1]], startCol = start.cols[i], startRow = start.rows[i], borders = "rows", sheet = sheet.names[i])
  openxlsx::addStyle(style = headerStyle, 
                     wb = wb, 
                     sheet = sheet.names[i], 
                     rows = start.rows[i], 
                     cols = (start.cols[i]:(length(report_list[i][[1]]) + (start.cols[i] - 1))), 
                     gridExpand = FALSE, 
                     stack = TRUE)
  
  openxlsx::addStyle(style = bodyStyle, 
                     wb = wb, 
                     sheet = sheet.names[i],
                     rows = c((start.rows[i]+1):(start.rows[i] + nrow(report_list[i][[1]]))), 
                     cols = c(start.cols[i]:(length(report_list[i][[1]]) + (start.cols[i] - 1))), 
                     gridExpand = TRUE)
}

# Insert Tabs
# addWorksheet(wb, sheet1.name, gridLines = FALSE)
# addWorksheet(wb, sheet2.name, gridLines = FALSE)

# Populate with data
# openxlsx::writeData(wb = wb, x = wk_report, startCol = 4, startRow = 4, borders = "rows", sheet = sheet1.name)
# openxlsx::writeData(wb = wb, x = issues, startCol = 1, startRow = 4, borders = "rows", sheet = sheet2.name)

# Format layouts and styles

setColWidths(wb, sheet1.name, cols = c(4, 5, 6, 7, 8, 9, 10, 11, 12), widths = c(10, 13, 13, 13, 13, 13, 13, 13, 13))

setColWidths(wb, sheet2.name, cols = c(1, 2, 3), widths = c(30, 55, 18))

#openxlsx::addStyle(wb, sheet = sheet1.name, style = headerStyle, rows = 4, cols = c(4:12), gridExpand = FALSE, stack = TRUE)
openxlsx::addStyle(wb, sheet = sheet1.name, style = highlight.row, rows = 6, cols = c(4:12), gridExpand = FALSE, stack = TRUE)
openxlsx::addStyle(wb, sheet = sheet1.name, style = underline.header, rows = 1, cols = c(1:15), gridExpand = FALSE, stack = TRUE)
openxlsx::addStyle(wb, sheet = sheet1.name, style = createStyle(numFmt = "##0%"), rows = c(5:9), cols = c(12), gridExpand = FALSE, stack = TRUE)
# openxlsx::addStyle(wb, sheet = sheet1.name, style = bodyStyle, rows = c(5:9), cols = c(4:12), gridExpand = TRUE, stack = TRUE)

# openxlsx::addStyle(wb, sheet = sheet2.name, style = bodyStyle.issues, rows = c(5:(4 + nrow(issues))), cols = c(1:3), gridExpand = TRUE)
openxlsx::addStyle(wb, sheet = sheet2.name, style = underline.header, rows = 1, cols = c(1:3), gridExpand = TRUE)
#openxlsx::addStyle(wb, sheet = sheet2.name, style = headerStyle.issues, rows = 4, cols = c(1:3), gridExpand = TRUE)

print(tasks_over_time)
insertPlot(wb, sheet = 1, xy = c("B", 11), width = 15, height = 10, fileType = "png", units = "cm")

print(late_versus_ontime)
insertPlot(wb, sheet = 1, xy = c("I", 11), width = 15, height = 10, fileType = "png", units = "cm")

if(nrow(completion_ratio) > 0) { # If there is enough data to calculate completion ratio otherwise skip
  print(completion_ratio_plt)
  insertPlot(wb, sheet = 1, xy = c("B", 31), width = 15, height = 9, fileType = "png", units = "cm")
}

print(plot_tasks_assigned_to)
insertPlot(wb, sheet = 1, xy = c("I", 31), width = 15, height = 9, fileType = "png", units = "cm")

# Format Summary page layout
openxlsx::pageSetup(wb, sheet = sheet1.name, orientation = "landscape", paperSize = 8)
openxlsx::pageSetup(wb, sheet = sheet2.name, orientation = "portrait", fitToWidth = TRUE)

# setHeader(wb, "TASK REPORT", position = "left")
setHeaderFooter(wb, sheet = sheet1.name, 
                header = c(paste0('&"Arial"&B&14&K008C98TASK REPORT - ',  
                                  ifelse(length(plan_name_to_filter) > 1, "", toupper(plan_name_to_filter))), 
                                                  NA, 
                                                  NA),
                footer = c("Printed On: &[Date]", NA, "Page &[Page] of &[Pages]"),)

setHeaderFooter(wb, sheet = sheet2.name, 
                header = c(paste0('&"Arial"&B&14&K008C98CURRENT ISSUES - ',  
                                  ifelse(length(plan_name_to_filter) > 1, "", toupper(plan_name_to_filter))), 
                           NA, 
                           NA),
                footer = c("Printed On: &[Date]", NA, "Page &[Page] of &[Pages]"),)


saveWorkbook(wb, here::here("output", paste0(Sys.Date(), " Task Report.xlsx")), overwrite = TRUE, returnValue = FALSE)

