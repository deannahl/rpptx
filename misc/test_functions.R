pres <- presentation(
  "../test_table_replacement/test_presentation.pptx"
)

new_table <- data.frame(
  column_01 = 1:4,
  column_02 = LETTERS[1:4],
  column_03 = round(runif(4), 2),
  column_04 = sample(colors(), size = 4)
)

replace_table(
  pres = pres,
  label = "test_table",
  new_table = new_table,
)

replace_image(
  pres = pres,
  label = "test_image",
  new_image = "../test_table_replacement/cute_dog.jpg"
)

replace_text(
  pres = pres,
  label = "test_text",
  new_text = "Imagine how hard physics would be if electrons had feelings.' Richard Feynmann. Welcome to our world, Dick."
)

data(iris)
library(tidyverse)
new_data <- iris %>%
  group_by(Species) %>%
  summarise("Sepal Width" = mean(Sepal.Width),
            "Sepal Length" = mean(Sepal.Length))

replace_category_plot(
  pres,
  label = "test_bar_plot",
  data = new_data
)

replace_donut_plot(
  pres,
  label = "test_donut_plot",
  percent_favorable = .86
)

save_pres(
  pres = pres,
  "../test_table_replacement/test_presentation_edit.pptx"
)
