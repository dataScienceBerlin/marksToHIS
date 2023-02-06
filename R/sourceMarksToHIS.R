getRange <- function(df, colnameMatrNrHIS = "Matrikelnummer", colnameNoteHIS = "Leistung"){
  rowHeader <- which(apply(df, MARGIN = 1, 
                           function(x) all(c(colnameMatrNrHIS, colnameNoteHIS) %in% x)))
  colMatrNrHIS <- which(df[rowHeader, ] == colnameMatrNrHIS)
  colNoteHIS <- which(df[rowHeader, ] == colnameNoteHIS)
  rowLastEntry <- min(which(is.na(df[(rowHeader + 1):nrow(df), colMatrNrHIS]))) + rowHeader - 1
  
  list(rows = (rowHeader + 1):rowLastEntry,
       cols = c(colMatrNrHIS, colNoteHIS))
}
  
getMarks <- function(df, ...){
  
  range <- getRange(df, ...)
  marksHis <- df[range$rows, range$cols] %>% 
    setNames(c("MatrNrHIS", "NoteHIS"))
  
  marksHis
}

joinAndReplaceMarks <- function(notenHIS, dfNoten, colMATRNRr, colMARKS) {
  notenHIS %>% 
    select(MatrNrHIS) %>% 
    left_join(dfNoten[ , c(colMATRNRr, colMARKS)] %>% mutate_all(as.character),
              by = c("MatrNrHIS" = colMATRNRr)) %>% 
    setNames(c("MatrNrHIS", "NoteHIS")) %>% 
    mutate(NoteHIS = ifelse(is.na(NoteHIS), "-", NoteHIS))
}

writeMarks <- function(df, notenNeu, ...) {
  
  # Check, dass sich Sortierung der Matr.-Nr nicht geändert hat, da Schreiben über
  # Positional-Matching
  
  stopifnot(all(getMarks(df)$MatrNrHIS == notenNeu$MatrNrHIS))
  
  # Ersetzen der Noten
  range <- getRange(df, ...)
  df[range$rows, range$cols] <- notenNeu
  
  df
}

