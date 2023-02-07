#' Identify range where IDs and marks are stored in HIS xlsx file
#'
#' @param df data.frame containing the HIS template
#' @param colIDs character column name where IDS (Matrikelnummer) are stored in template
#' @param colMarks character column name where Marks (Noten) are stored in template
getRange <- function(df, colIDs = "Matrikelnummer", colMarks = "Leistung"){
  rowHeader <- which(apply(df, MARGIN = 1,
                           function(x) all(c(colIDs, colMarks) %in% x)))
  colMatrNrHIS <- which(df[rowHeader, ] == colIDs)
  colNoteHIS <- which(df[rowHeader, ] == colMarks)
  rowLastEntry <- min(which(is.na(df[(rowHeader + 1):nrow(df), colMatrNrHIS]))) + rowHeader - 1

  list(rows = (rowHeader + 1):rowLastEntry,
       cols = c(colMatrNrHIS, colNoteHIS))
}

#' Extract Marks from HIS xlsx file
#'
#' @inheritParams getRange
#' @param ... further arguements passed to `getRange`
#'
#' @return data.frame with columns `ID` and `mark`
#'
#' @export
marksFromHISxlsx <- function(df, ...){

  range <- getRange(df, ...)
  marksHis <- df[range$rows, range$cols]
  names(marksHis) <- c("ID", "mark")

  marksHis
}

#' Replace Marks with existing ones based on ID mapping
#'
#' @param marksHIS data.frame marks extracted from HIS xlsx template
#' @param dfMarks data.frame with existing marks to be inserted into `marksHIS`
#' @param colID character column name containing the IDs in `marksHIS`
#' @param colMarks character column name containing the marks in `marksHIS`
#' @param strNotAttended character indicating HIS that student did not attend the exam
#'
#' @details Currently marks for `ID`s in `marksHIS` that aren't contained in `dfMarks`
#' are set to `strNotAttended` which is interpreted by HIS as *not attended*.
#'
#' @returns updated marks (data.frame with columns `ID` and `mark`)
#'
#'
#' @export
replaceMarks <- function(marksHIS, dfMarks, colID, colMarks, strNotAttended = "-") {
  df <- marksHIS %>%
    select(matches("ID")) %>%
    left_join(dfMarks[ , c(colID, colMarks)] %>% mutate_all(as.character),
              by = c("ID" = colID)) %>%
    setNames(c("ID", "mark"))
  if(any(is.na(df$mark))){
    df$mark[is.na(df$mark)] <- strNotAttended
  }

  # warning, if new IDs are contained in `dfMarks`
  if(!(all(dfMarks[[colID]] %in% marksHIS$ID))) {
    IDsmissing <- dfMarks[[colID]][!dfMarks[[colID]] %in% marksHIS$ID]
    warning("The following IDs are not contained marksHIS:\n",
            paste0(paste0(" - ", IDsmissing), collapse = "\n"))
  }

  df

}

#' Writes tabular ID-marks- data to HIS xlsx template
#' @inheritParams getRange
#' @inheritParams marksFromHISxlsx
#' @param marksUpdated data.frame tabular data with columns `ID` and `mark`
#'
#' @returns updated HIS xlsx file
#'
#' @export
#'
#' @examples
#' \dontrun{
#' # read xlsx file exported from HIS
#' require(openxlsx)
#' dfHIS <-  read.xlsx(xlsxFile = "HIS_xlsx_template.xlsx", colNames = FALSE,
#'                     skipEmptyRows = FALSE, skipEmptyCols = FALSE)
#' # read file with existing marks
#' # with colums 'ID' and 'mark'
#' dfMarks <- read.csv(file = "marks.csv")
#'
#' # Workflow to update marks
#' marksHIS <- marksFromHISxlsx(dfHIS)
#' marksNew <- replaceMarks(marksHIS = marksHIS, dfMarks = dfMarks,
#'                          colID = "ID", colMarks = "mark")
#' dfHISnew <- marksToHISxlsx(df = dfHIS, marksUpdated = marksNew)
#'
#' # write updated xlsx file to be imported to HIS
#' write.xlsx(dfHISnew, file = "HIS_updated.xlsx", colNames = FALSE, rowNames = FALSE)
#' }
marksToHISxlsx <- function(df, marksUpdated, ...) {

  # Check, dass sich Sortierung der Matr.-Nr nicht geändert hat, da Schreiben über
  # Positional-Matching

  stopifnot(all(marksFromHISxlsx(df)$ID == marksUpdated$ID))

  # Ersetzen der Noten
  range <- getRange(df, ...)
  df[range$rows, range$cols] <- marksUpdated

  df
}

