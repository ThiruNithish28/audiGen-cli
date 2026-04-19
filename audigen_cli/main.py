from audigen_cli.llm_client import callLLM
from audigen_cli.extractor import extractDoc
from audigen_cli.excelWriter import startExcelChange


text = extractDoc();
response = callLLM(text);
startExcelChange(response,"20-4-2024", "30-4-2024")
print("final" + "\n" + "done");
