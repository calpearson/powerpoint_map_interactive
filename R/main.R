# Load libraries
library(dplyr)
library(ggplot2)
library(leaflet)
library(htmlwidgets)
library(officer)
library(rnaturalearth)
library(rnaturalearthdata)
library(sf)
library(gridExtra)
library(grid)

# Ensure output folder exists
if(!dir.exists("output")) dir.create("output")

# 1. Dummy outbreak dataset
epi_data <- data.frame(
  country = c("Vietnam", "Thailand", "Malaysia", "Singapore", "Philippines", "Indonesia","Laos","Cambodia"),
  cases = sample(500:5000, 8),
  deaths = sample(10:300, 8)
) %>%
  mutate(cfr = round((deaths / cases) * 100, 2))

# 2. SE Asia map
se_asia <- ne_countries(scale = "medium", returnclass = "sf") %>%
  filter(admin %in% epi_data$country)

map_data <- se_asia %>%
  left_join(epi_data, by = c("admin" = "country"))

# 3. Static map
p <- ggplot(map_data) +
  geom_sf(aes(fill = cases), color = "black") +
  scale_fill_gradient(low = "orange", high = "red") +
  labs(title = "Southeast Asia Flu Outbreak", fill = "Cases") +
  theme_minimal()

ggsave("output/static_map.png", plot = p, width = 7, height = 5)

# 4. Interactive map (self-contained)
# m <- leaflet(map_data) %>%
#   addTiles() %>%
#   addPolygons(
#     fillColor = ~colorNumeric("YlOrRd", cases)(cases),
#     weight = 1,
#     color = "black",
#     fillOpacity = 0.7,
#     popup = ~paste0(admin, "<br>Cases: ", cases, "<br>Deaths: ", deaths, "<br>CFR: ", cfr, "%")
#   )

m <- leaflet(map_data) %>%
  addTiles() %>%
  addPolygons(
    fillColor = ~colorNumeric("YlOrRd", cases)(cases),
    weight = 1,
    color = "black",
    fillOpacity = 0.7,
    popup = ~paste0(
      "<b>", admin, "</b><br>",
      "Cases: ", cases, "<br>",
      "Deaths: ", deaths, "<br>",
      "CFR: ", cfr, "%"
    )
  )

# Save as self-contained HTML
saveWidget(m, "output/interactive_map.html", selfcontained = TRUE)


# Self-contained HTML so PowerPoint can open it directly
saveWidget(m, "output/interactive_map.html", selfcontained = TRUE)

# 5. Create tight epi table
tbl <- tableGrob(epi_data, rows = NULL)
tbl$heights <- unit(rep(1.5, nrow(epi_data) + 1), "lines")
tbl$widths <- unit(rep(1, ncol(epi_data)), "null")

png("output/epi_table.png", width = 900, height = 250, res = 100)
grid.draw(tbl)
dev.off()

# 6. Hyperlink text object (points to self-contained HTML)
link_txt <- fpar(
  hyperlink_ftext("Open Interactive Map", "interactive_map.html",
                  prop = fp_text(color = "blue", underline = TRUE, font.size = 16))
)

# 7. Create PowerPoint with 3 slides
ppt <- read_pptx() %>%
  
  # Slide 1: Title
  add_slide(layout = "Title Slide", master = "Office Theme") %>%
  ph_with(value = "Flu Outbreak in Southeast Asia", location = ph_location_type(type = "ctrTitle")) %>%
  ph_with(value = "Dummy Data Example", location = ph_location_type(type = "subTitle")) %>%
  
  # Slide 2: Overview (epi table)
  add_slide(layout = "Title and Content", master = "Office Theme") %>%
  ph_with(value = "Overview of Cases and Mortality", location = ph_location_type(type = "title")) %>%
  ph_with(external_img("output/epi_table.png"), location = ph_location(left = 1, top = 1.5, width = 8, height = 2.5)) %>%
  
  # Slide 3: Map + hyperlink
  add_slide(layout = "Title and Content", master = "Office Theme") %>%
  ph_with(value = "Flu Outbreak Map", location = ph_location_type(type = "title")) %>%
  ph_with(external_img("output/static_map.png"), location = ph_location(left = 0.5, top = 1.5, width = 9, height = 5)) %>%
  ph_with(link_txt, location = ph_location(left = 3, top = 6.2, width = 3, height = 0.5))

# 8. Save PowerPoint in output folder
print(ppt, target = "output/flu_outbreak_presentation.pptx")
