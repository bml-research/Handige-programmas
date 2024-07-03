library(rstudioapi)

## read text files
work_path <- setwd(dirname(rstudioapi::getActiveDocumentContext()$path))
#sub.folders <- list.dirs(work_path, recursive = T)[-1]
new.folder <- paste0(work_path, "/new")
list_of_files <- list.files(work_path, "*.zip")
#list_of_files2 <- list.files(new.folder)
file.copy(file.path(work_path, list_of_files), new.folder)

for(i in 1:length(list_of_files)) {
  #file.copy(file.path(work_path, list_of_files[i]), new.folder)
  file.rename(list_of_files[i], paste0("dataset00", i + 40, ".zip"))

    
}
