from doctr.io import DocumentFile
from doctr.models import ocr_predictor

doc = DocumentFile.from_pdf("/content/Justworks - Empower Semiconductor, Inc. Invoice Details Report 02_06_26.pdf")
model = ocr_predictor(pretrained=True)
result = model(doc)

# Save extracted text
with open("output.txt", "w") as f:
    f.write(result.render())