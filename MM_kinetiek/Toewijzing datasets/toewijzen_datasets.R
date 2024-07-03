library(openxlsx)
library(xlsx)
library(tools)

read_path <- "/Users/markuiz/PycharmProjects/Eigen projectjes/MM_kinetiek/Toewijzing datasets/"
df <- openxlsx::read.xlsx(paste0(read_path,"presentielijst_2023-2024.xlsx"), 1, cols = 1)
geordend <- df[order(df$Studentnummer),]
datasetjes <- c()
for(i in 1:nrow(df)) {
  plakken <- paste0("dataset", i)
  datasetjes <- append(datasetjes, plakken)
}

random_datasetjes <- sample(datasetjes)
verdeling_plots <- sample(c("Lineweaver-Burke", "Eadie-Hofstee", "Hanes-Woolf"), 88,
                replace = TRUE)

new_df <- cbind(geordend, random_datasetjes)
new_df <- cbind(geordend, random_datasetjes, verdeling_plots)

colnames(new_df) <- c("Studentnummer", "Dataset")
colnames(new_df) <- c("Studentnummer", "Dataset", "Plot")

write.xlsx2(new_df, paste0(read_path, "verdeling_datasets_kinetiek.xlsx"), col.names = TRUE,
            row.names = FALSE)
