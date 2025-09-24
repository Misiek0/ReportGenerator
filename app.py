import word as w
import excel as x


word_stationary = w.open_docx("Raport odnośnie urządzeń umieszczonych w terenie.docx")
word_mobile = w.open_docx("Raport odnośnie urządzeń mobilnych umieszczonych w pociągach.docx")

table_stationary = word_stationary.tables[0]
w.find_col_name("Usuwanie nieprawidłowości w funkcjonowaniu urządzeń",table_stationary)

table_mobile = word_mobile.tables[0]
w.find_col_name("Usuwanie nieprawidłowości w funkcjonowaniu urządzeń",table_mobile)





