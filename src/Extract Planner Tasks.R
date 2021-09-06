#remotes::install_github("giocomai/ganttrify", dependencies = TRUE)
library(ganttrify)
# data\MS Planner Exports
library(openxlsx)
library(ggplot2)
library(dplyr)

# Parameters --------------------------------------------------------------

task_week_view_ahead <- 1 # Number of week(s) beyond and behind current week
task_week_view_behind <- 0 # Number of week(s) behind current week

plan_name_to_filter <- c("PID840 – Spoilbank Marina PMG", "Ideas Hub") # Used to filter report
# plan_name_to_filter <- c("PID840 – Spoilbank Marina PMG") # Used to filter report

# Import Data -------------------------------------------------------------

data_path <- here::here("data", "MS Planner Exports")

files <- list.files(data_path, pattern = "*.xlsx", full.names = TRUE)

temp <- NULL

for(i in seq_along(files)){
  
  xl_file <- readxl::read_xlsx(path = files[i])
  
  file.info(files[i])$mtime
  
  plan_id <- as.character(xl_file[1, 2])
  
  plan_name <- names(xl_file)[2]
  
  # Extract the file modification date/time from the file info:
  planner_version <- file.info(files[i])$mtime
  
  # Extract the date from the planner file spreadsheet
  # planner_version <- lubridate::parse_date_time(as.character(xl_file[2, 2]), 
  #                                               orders = c("mdY"))
  
  col_names <- make.names(xl_file[4, ])
  
  xl_file <- xl_file[-1:-4, ]
  
  xl_file$plan_version <- planner_version
  
  xl_file$plan_id <- plan_id
  
  xl_file$Plan.Name <- plan_name
  
  names(xl_file) <- c(col_names, "Plan.Version", "Plan.ID", "Plan.Name")
  
  temp <- rbind(temp, xl_file)
  
}

plan.names <- data.frame(Plan.ID = unique(temp$Plan.ID),
                         Plan.Name = unique(temp$Plan.Name))

temp <- temp %>% mutate(Week.Number.Version = strftime(Plan.Version, format = "%Y-W%V"),
                        Completed.Date = lubridate::parse_date_time(Completed.Date, orders = c("mdY")),
                        Week.Number.Completed = strftime(Completed.Date, format = "%Y-W%V"),
                        Created.Date = lubridate::parse_date_time(Created.Date, orders = c("mdY")),
                        Start.Date = lubridate::parse_date_time(Start.Date, orders = c("mdY")),
                        Due.Date = lubridate::parse_date_time(Due.Date, orders = c("mdY")),
                        Week.Number.Due = strftime(Due.Date, format = "%Y-W%V"),
                        Start.Date = case_when(is.na(Start.Date) ~ Created.Date,
                                               TRUE ~ Start.Date))

all_planner_tasks <- temp

# Tasks Current Status ----------------------------------------------------

tasks_current_status <- all_planner_tasks %>%
  select(Task.ID, Task.Name, Plan.Version) %>%
  distinct(Task.ID, Task.Name, .keep_all = TRUE) %>%
  group_by(Task.ID) %>%
  mutate(newest = max(Plan.Version)) %>%
  filter(Plan.Version == newest) %>%
  left_join(temp %>%
              select(Plan.ID,
                     Task.ID,
                     Plan.Name,
                     Bucket.Name,
                     Plan.Version,
                     Created.By,
                     Created.Date,
                     Progress,
                     Start.Date,
                     Due.Date,
                     Week.Number.Due,
                     Late),
            by = c("Task.ID" = "Task.ID",
                   "Plan.Version" = "Plan.Version")) %>%
  select(Plan.ID,
         Task.ID,
         Plan.Name,
         Bucket.Name,
         Plan.Version,
         Task.Name,
         Created.By,
         Created.Date,
         Progress,
         Start.Date,
         Due.Date,
         Week.Number.Due,
         Late)

# How Many Times Have Due Dates Been Pushed Back? -------------------------

latest_plan_version <- all_planner_tasks %>% 
  select(Plan.ID, Plan.Version,Plan.Name) %>% 
  group_by(Plan.ID) %>% 
  filter(Plan.Version == max(Plan.Version)) %>% 
  ungroup() %>% 
  distinct()

latest_plan <- all_planner_tasks %>% 
  group_by(Plan.ID) %>% 
  filter(Plan.Version == max(Plan.Version)) %>% 
   ungroup()

# Look up the Plan.ID based on the plan_name_to_filter values entered in parameters:
selected_plan <- as_tibble(latest_plan_version[grep(paste(plan_name_to_filter, collapse="|"), 
                                          latest_plan_version$Plan.Name), 1])

selected_plan <- as_tibble(latest_plan_version[grep(paste(plan_name_to_filter, collapse="|"), 
                                                    latest_plan_version$Plan.Name), 1])

unique_tasks <- all_planner_tasks %>% select(Task.ID) %>% 
  distinct()

changed_due_dates <- all_planner_tasks %>% 
  filter(Plan.ID %in% selected_plan$Plan.ID) %>% 
  select(Task.ID, Due.Date) %>% 
  distinct(.keep_all = TRUE) %>%
  group_by(Task.ID) %>% 
  count() %>% 
  arrange(desc(n)) %>% 
  rename(due_date_changes = n) %>% 
  left_join(latest_plan %>% 
              filter(Plan.ID %in% selected_plan) %>% 
              select(Task.ID, Task.Name, Progress, Due.Date, Late))

# Tasks by Person ---------------------------------------------------------

tasks_by_person <- tidyr::separate_rows(data = latest_plan, Assigned.To, sep = ";")

# Status Over Time --------------------------------------------------------

all_planner_tasks %>%
  select(Plan.ID, Plan.Name, Week.Number.Version, Plan.Version, Progress) %>% 
  group_by(Plan.ID, Week.Number.Version, Plan.Name) %>% 
  mutate(newest = max(Plan.Version)) %>% 
  filter(Plan.Version == newest,
        # Plan.Version == max(Plan.Version),
         Plan.ID %in% selected_plan$Plan.ID) %>%
  ggplot(aes(x = Week.Number.Version, fill = Progress)) +
  geom_bar() +
  facet_wrap(~factor(Plan.Name))

# Task KPIs ---------------------------------------------------------------

filtered_latest_plan <- latest_plan %>% 
  filter(Plan.ID %in% selected_plan$Plan.ID)

min_start_date <- min(c(min(filtered_latest_plan$Due.Date), 
                        min(filtered_latest_plan$Created.Date),
                        min(filtered_latest_plan$Completed.Date)), na.rm = TRUE)

max_end_date <- max(c(max(filtered_latest_plan$Due.Date), 
                      max(filtered_latest_plan$Created.Date), 
                      max(filtered_latest_plan$Completed.Date)), na.rm = TRUE)
  
weeks <- data.frame(Week.Number.Completed = strftime(seq(as.Date(min_start_date), as.Date(max_end_date), by = "week"), format = "%Y-W%V"))

current_week <- strftime(Sys.Date(), format = "%Y-W%V")
last_week <- paste0(lubridate::year(Sys.Date()), "-W", lubridate::isoweek(Sys.Date()) - task_week_view_behind)
next_week <- paste0(lubridate::year(Sys.Date()), "-W", lubridate::isoweek(Sys.Date()) + task_week_view_ahead)

planned_tasks_due <- filtered_latest_plan %>% 
  group_by(Week.Number.Due, Plan.ID) %>% 
 # filter(Progress != "Completed") %>% 
  count() %>% 
  rename(planned.tasks.due = n,
         Week.Number.Completed = Week.Number.Due) # This enables joining later

planned_tasks_complete_ontime <- filtered_latest_plan %>% 
  filter(Progress == "Completed",
         Week.Number.Due == Week.Number.Completed # & Week.Number.Version == Week.Number.Completed 
         ) %>% 
  group_by(Week.Number.Completed, Plan.ID) %>% 
  count() %>% 
  rename(planned.tasks.complete.ontime = n)

unplanned_tasks_complete <- filtered_latest_plan %>% 
  filter(grepl("unplanned task", 
               tolower(filtered_latest_plan$Labels)) == TRUE,
         Progress == "Completed"
         ) %>% 
  group_by(Week.Number.Completed, Plan.ID) %>% 
  count() %>% 
  rename(unplanned.tasks.completed = n)

planned_tasks_complete_late <- filtered_latest_plan %>% 
  filter(Progress == "Completed",
         Week.Number.Due < Week.Number.Completed) %>% 
  group_by(Week.Number.Completed, Plan.ID) %>% 
  count() %>% 
  rename(planned.tasks.completed.late = n)

planned_tasks_complete_early <- filtered_latest_plan %>% 
  filter(Progress == "Completed",
         Week.Number.Due > Week.Number.Completed) %>% 
  group_by(Week.Number.Completed, Plan.ID) %>% 
  count() %>% 
  rename(planned.tasks.completed.early = n)

tasks_kpis <- weeks %>% 
  left_join(planned_tasks_due) %>% 
  left_join(planned_tasks_complete_ontime) %>% 
  left_join(unplanned_tasks_complete) %>% 
  left_join(planned_tasks_complete_late) %>% 
  left_join(planned_tasks_complete_early) %>% 
  rename(Week.Number = Week.Number.Completed)

tasks_by_week <- tasks_kpis %>% 
  rowwise() %>% 
  mutate(Total.Tasks.Complete = sum(c_across(planned.tasks.complete.ontime:planned.tasks.completed.early))) %>% 
  tidyr::replace_na(list(planned.tasks.due = 0,
                         planned.tasks.complete.ontime = 0,
                         unplanned.tasks.completed = 0,
                         planned.tasks.completed.late = 0,
                         planned.tasks.completed.early = 0,
                         Total.Tasks.Complete = 0))

current_week_report <- tasks_by_week  %>% 
  filter(Week.Number >= last_week &
           #Week.Number == current_week |
           Week.Number <= next_week)

tasks_by_week_pivot <- tasks_by_week %>% 
  select(!c(planned.tasks.due, Total.Tasks.Complete)) %>% 
  tidyr::pivot_longer(!Week.Number:Plan.ID, 
                      names_to = "Task.Status", 
                      values_to = "count")


# Plots -------------------------------------------------------------------

plotly::ggplotly(tasks_by_week_pivot %>%
  filter(Week.Number <= next_week) %>% 
  ggplot(aes(x = Week.Number, 
             y = count, 
             fill = Task.Status)) +
  geom_col() +
  theme_bw())

# GANTT Chart -------------------------------------------------------------

dat_gant <- latest_plan %>% ungroup() %>% 
  filter(Progress != "Completed", 
         Due.Date >= as.Date(latest_plan_version$Plan.Version)) %>% 
  select(Bucket.Name, Task.Name, Start.Date, Due.Date, Progress) %>% 
  rename(wp = Bucket.Name, 
         activity = Task.Name, 
         start_date = Start.Date, 
         end_date = Due.Date,
         spot_type = Progress) %>% 
  mutate(start_date = format(as.Date(start_date), "%Y-%m-%d"),
         end_date = format(as.Date(end_date), "%Y-%m-%d"))

  ganttrify(project = dat_gant, 
            by_date = TRUE,
            hide_wp = F,
            spots = dat_gant$spot_type,
            size_text_relative = 1, 
            axis_text_align = "left"
            # font_family = "Roboto Condensed",
          #  project_start_date = "2021-01-01"
            ) +
    ggplot2::labs(title = plan_name,
                  subtitle = paste0("Plan Version Date: ", 
                                    format(as.Date(latest_plan_version$Plan.Version), 
                                           "%d %B %Y")),
                  caption = "") +
    ggplot2::geom_vline(xintercept = as.Date(Sys.Date()), 
                        linetype = "dashed", 
               color = "black", size = 1) +
    coord_cartesian(xlim = c(as.Date(Sys.Date()), 
                           max(as.Date(dat_gant$end_date))))

  