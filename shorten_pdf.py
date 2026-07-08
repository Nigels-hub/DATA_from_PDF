import pypdfium2 as pdfium

src = pdfium.PdfDocument('2024_10_30_Modul4A_Zolbetuximab_Anhang_4_G4.pdf')
dst = pdfium.PdfDocument.new()
dst.import_pages(src, list(range(min(100, len(src)))))
dst.save('TEST_100page.pdf')
print(f'Done — TEST1.pdf created with {len(dst)} pages')
