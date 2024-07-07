import json

def extract_json_from_text(text:str, prefix=None):
   text=text.replace("```json","")
   return json.loads(text.replace("```",""))