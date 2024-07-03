library(openxlsx)
library(xlsx)
library(tools)
library(rstudioapi)

# location for reading and writing data
work_path <- setwd(dirname(rstudioapi::getActiveDocumentContext()$path))

df <- read.xlsx(paste0(work_path,"/presentielijst.xlsx"), 1, colIndex = 1)
geordend <- df[order(df$StudentNumber),]
datasetjes <- c()
for(i in 1:nrow(df)) {
  plakken <- paste0("dataset", i)
  datasetjes <- append(datasetjes, plakken)
}

random_datasetjes <- sample(datasetjes)
new_df <- cbind(geordend, random_datasetjes)

colnames(new_df) <- c("Studentnummer", "Dataset")

write.xlsx2(new_df, paste0(work_path, "/verdeling_datasets.xlsx"), col.names = TRUE,
            row.names = FALSE)
