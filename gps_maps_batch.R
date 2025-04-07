#### gps data plotting - batch processing ####
# works, but still having issues with the end point labeling 

library(tidyverse)
library(leaflet)
library(readxl)
library(htmlwidgets)
library(webshot)
library(beepr)

# process each sheet in an Excel file
process_sheet <- function(file_path, sheet_name, output_map_name, sheet_index) {
  
  # read the sheet to get the number of rows in the data
  sheet_data <- read_xlsx(file_path, sheet = sheet_name, 
                          range = "A3:D10000", col_names = FALSE) %>%
    set_names(c("lat", "lon", "date_time", "location"))
  
  # find the last row with data (ignoring empty rows)
  last_row <- which(!is.na(sheet_data$lat))[length(which(!is.na(sheet_data$lat)))]  # Last row with lat data
  
  # adjust range dynamically
  range <- paste0("A3:D", last_row)  
  
  # read in the data based on the adjusted range
  pl_1 <- read_xlsx(file_path, sheet = sheet_name, range = range, col_names = FALSE) %>%
    set_names(c("lat", "lon", "date_time", "location")) %>%
    slice(-1)  
  
  # convert date_time to proper datetime format
  pl_1 <- pl_1 %>%
    mutate(
      date_time = as_datetime((as.numeric(date_time) - 25569) * 86400),
      lat = as.numeric(lat),
      lon = as.numeric(lon)
    ) %>%
    filter(!is.na(lat), !is.na(lon))  
  
  # Time difference and movement status 
  pl_1 <- pl_1 %>%
    arrange(date_time) %>%
    mutate(
      time_diff = c(NA, difftime(date_time[-1], date_time[-nrow(pl_1)], units = "mins"))
    ) %>%
    mutate(
      movement_status = case_when(
        time_diff > 10 & lag(time_diff, default = 0) <= 10 ~ "start",
        time_diff <= 10 & lag(time_diff, default = 0) > 10 ~ "stop",
        TRUE ~ "none"
      )
    )
  
  # Filter rows where movement starts or stops
  pl_1_label <- pl_1 %>% filter(movement_status %in% c("start", "stop"))
  
  # Calculate lat/lon bounds
  lat_min <- min(pl_1$lat, na.rm = TRUE)
  lat_max <- max(pl_1$lat, na.rm = TRUE)
  lon_min <- min(pl_1$lon, na.rm = TRUE)
  lon_max <- max(pl_1$lon, na.rm = TRUE)
  
  # Choose every nth point for readability
  n <- 20
  pl_1_label_on_line <- pl_1[seq(1, nrow(pl_1), by = n), ]
  
  # Access start and end points
  start_point <- pl_1[1, ]
  end_point <- tail(pl_1, 1)
  
  # Debug: Print last few rows to verify data
  print(paste("Processing sheet:", sheet_name))
  print(head(pl_1, 5))  # Print first few rows to check
  
  # Create dynamic names
  df_name <- paste0("pl_", sheet_index)
  map_name <- paste0("pl_", sheet_index, "_map")  # Dynamic map name
  
  # Create the map
  pl_1_map <- leaflet(pl_1) %>%
    addTiles(urlTemplate = "https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}",
             attribution = "ESRI World Imagery") %>%
    
    # Path (line) connecting GPS points
    addPolylines(data = pl_1, lat = ~lat, lng = ~lon, color = "red", weight = 2, opacity = 1) %>%
    
    # Labels along the line
    addLabelOnlyMarkers(data = pl_1_label_on_line, lat = ~lat, lng = ~lon,
                        label = ~paste(format(date_time, "%Y-%m-%d %H:%M:%S")),
                        labelOptions = labelOptions(noHide = TRUE, direction = 'top', textsize = '10px', style = list("color" = "blue", "font-weight" = "bold"))) %>%
    
    # Start and end location labels
    addLabelOnlyMarkers(data = start_point, lat = ~lat, lng = ~lon,
                        label = ~paste("Start Location:", location),
                        labelOptions = labelOptions(noHide = TRUE, direction = 'top', textsize = '10px', style = list("color" = "green", "font-weight" = "bold"))) %>%
    
    addLabelOnlyMarkers(data = end_point, lat = ~lat, lng = ~lon,
                        label = ~paste("End Location:", location),
                        labelOptions = labelOptions(noHide = TRUE, direction = 'top', textsize = '10px', style = list("color" = "red", "font-weight" = "bold"))) %>%
    
    # Legend for start/stop 
    addLegend("bottomright", 
              colors = c("green", "red"), 
              labels = c(paste("start location:", start_point$location),
                         paste("end location:", end_point$location)),
              title = paste("trip from", start_point$location, "to", end_point$location)) %>%
    
    # Adjust zoom level and view using calculated lat/lon bounds
    setView(lng = (lon_min + lon_max) / 2, lat = (lat_min + lat_max) / 2, zoom = 12)
  
  # Save map as HTML
  saveWidget(pl_1_map, output_map_name, selfcontained = TRUE)
  
  # Debug: Print the name of the map and df being saved
  print(paste("Saving map:", output_map_name))
  
  # Assign the map and data frame to global environment
  assign(map_name, pl_1_map, envir = .GlobalEnv)
  assign(df_name, pl_1, envir = .GlobalEnv)
}

# Example loop for processing files and sheets
file_dir <- getwd()
xlsx_files <- list.files(file_dir, pattern = "\\.xlsx$", full.names = TRUE)

for (file in xlsx_files) {
  sheet_names <- excel_sheets(file)
  
  for (i in seq_along(sheet_names)) {
    output_map_name <- paste0(tools::file_path_sans_ext(basename(file)), "_", i, "_map.html")
    process_sheet(file, sheet_names[i], output_map_name, i)
  }
}
