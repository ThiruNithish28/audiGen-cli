import time

from google import genai
from google.genai import types
from google.genai.errors import APIError
from pydantic import BaseModel
from typing import List
import os
from dotenv import load_dotenv
from rich.console import Console

console = Console()

# Environment configuration
load_dotenv()
api_key = os.getenv("GEMINI_API_KEY")

# Initialize client globally once
client = genai.Client(api_key=api_key)

# Data structures 
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


def get_response_from_model(prompt: str, status: any = None) : 
  """
  FallBack calls , for gemini model with free tire , to handle the rate limit
  
  @param prompt -> a final prompt send to LLM model
  @param status -> which is used to update the text prin in the console based on fallback  
  """
  
  models = ["gemini-2.5-pro","gemini-2.5-flash", "gemini-2.5-flash-lite"]
  
  for model in models:
    try:
      
      # Update status text if a fallback model is running
      if status and model != "gemini-2.5-pro":
        status.update(f"[bold cyan][2/3] Generating with fallback ({model})....[/bold cyan]")
        
      response = client.models.generate_content(
          model=model,
          contents=prompt,
          config=types.GenerateContentConfig(
            response_mime_type='application/json',
            response_schema=ResponseFormat,
            temperature=0.1,  # Low temperature ensures strict schema adherence
        )
      )
      return response
    except APIError as error:
      if error.code == 429:
        if status:
          status.update(f"\n  [yellow]⚠ Rate limit hit on {model}. Retrying with fallback...[/yellow]")
        else:
          console.print(f"\n  [yellow]⚠ Rate limit hit on {model}. Retrying in 2s...[/yellow]")
        time.sleep(2)
        continue
      else:
        console.print(f"\n  [red]🗙 API Error on {model}: {error.message}[/red]")
        continue
        
  # if all model fails  
  return None
      
# interact with gemini model to generate test case  
def callLLM(sanitized_text: str, status: any=None) :
  prompt = f"""
    Generate exactly 13 test cases for the following BRD.

    Also extract the "components affected" and "rasied by" EXACTLY as written in the BRD.
    the "components affected" will usually appear near phrases like: Requirement Widget / Module


    Return them strictly in the given JSON format.

    BRD:
    {sanitized_text}
  """

  response = get_response_from_model(prompt, status=status)
  
  if response and response.parsed:
    llm_generateTestCase=response.parsed
    return llm_generateTestCase
  
  console.print(f"\n[bold red]🗙 Critical: All Gemini models failed to return a response[/bold red]")
  return None