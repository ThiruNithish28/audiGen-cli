from google import genai
from pydantic import BaseModel
from typing import List
import os
from dotenv import load_dotenv

load_dotenv()

api_key = os.getenv("GEMINI_API_KEY")


class TestCase(BaseModel):
  testCaseId: int
  testCase: str
  expectedResult: str

class AdditionalInfo(BaseModel):
  componets_affected: str
  brd_rasiedBy: str

class ResponseFormat(BaseModel):
  testCases: List[TestCase]
  additional_info: AdditionalInfo

def callLLM(sanitized_text) :
  prompt = f"""
    Generate exactly 13 test cases for the following BRD.

    Also extract the "components affected" and "rasied by" EXACTLY as written in the BRD.
    the "components affected" will usually appear near phrases like: Requirement Widget / Module


    Return them strictly in the given JSON format.

    BRD:
    {sanitized_text}
  """

  client = genai.Client(api_key=api_key)

  response = client.models.generate_content(
      model="gemini-2.5-flash",
      contents=prompt,
      config=genai.types.GenerateContentConfig(
          response_mime_type= 'application/json',
          response_schema=ResponseFormat,

      )
  )
  print(response.text)
  llm_generateTestCase=response.parsed
  return llm_generateTestCase