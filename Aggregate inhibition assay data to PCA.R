library(rio)
library(xlsx)

input_path= "D:/input_data_plate_reader"
filepath <- list.files(input_path, pattern = ".xlsx$", full.names = TRUE)

File=NULL
for (i in 1:length(filepath)) {
  dataset <- read_excel(filepath[i], sheet = "Sheet 1")
  File= rbind(File, dataset)
}

output_pth= file.path(input_path, paste("PCA_inhibition_assay", ".xlsx", sep=""))
export(File, output_pth)