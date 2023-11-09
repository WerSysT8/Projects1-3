from publication_by_hour import generate_Kolichestvo_publicasi
from number_of_line_on_duty import generate_Kolichestvo_strok
from AI_filing_correctness import generate_Tabl_Korekt_zapoln_neiron
from create_PDF import generate_pdf
from tgBOT import send_file


if __name__ == "__main__":
    generate_Kolichestvo_publicasi()
    generate_Kolichestvo_strok()
    generate_Tabl_Korekt_zapoln_neiron()
    generate_pdf()
    # send_file()
