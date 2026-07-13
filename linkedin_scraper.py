
import requests
import json
import pandas as pd
import time
import os
from google import genai

# ---------------------------------------------------------
# 1. Configuration & API Keys
# ---------------------------------------------------------
SERPER_API_KEY = "YOUR_SERPER_KEY"
GEMINI_API_KEY = "YOUR_GEMINI_KEY"

os.environ["GEMINI_API_KEY"] = GEMINI_API_KEY
client = genai.Client()

# ---------------------------------------------------------
# 2. Serper Search Function
# ---------------------------------------------------------
def get_linkedin_profiles(sector, company, job_title, city, page_num):
    url = "https://google.serper.dev/search"
    query = f"site:linkedin.com/in/ {job_title} {company} {sector} {city}"
    
    payload = json.dumps({
        "q": query,
        "num": 10,
        "page": int(page_num)
    })
    
    headers = {
        'X-API-KEY': SERPER_API_KEY,
        'Content-Type': 'application/json'
    }
    
    try:
        response = requests.request("POST", url, headers=headers, data=payload)
        if response.status_code == 200:
            return response.json().get("organic", [])
        else:
            print(f"Error calling Serper: {response.text}")
            return []
    except Exception as e:
        print(f"Network error calling Serper: {e}")
        return []

# ---------------------------------------------------------
# 3. Combined Gemini Brain 
# ---------------------------------------------------------
def generate_lead_data(title, snippet):
    prompt = f"""
    You are an expert sales assistant. Analyze the following Google search result and extract/generate four specific pieces of information.
    
    1. Name: Extract the person's full name.
    2. Company: Extract the name of the company they are currently working for.
    3. Position: Extract their current job title/position.
    4. Pitch: Write a friendly, professional 600-700 character LinkedIn connection note based on these rules:
       - 1 sentence praising their current role.
       - Exclude past jobs/education.
       - Pitch "Sipwise Smart Bottles" (briefly explain product & B2B collab value).
       - End by asking for a quick online meet.

    Input Data:
    Title: {title}
    Snippet: {snippet}

    CRITICAL INSTRUCTION: You must separate your output for Name, Company, Position, and Pitch using exactly '|||' as the delimiter. Do not include any other markdown, labels, or headers (like 'Name:' or 'Pitch:').
    
    Expected Output Format:
    [Name] ||| [Company] ||| [Position] ||| [Your connection note pitch here]
    """
    
    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt
    )
    return response.text

# ---------------------------------------------------------
# 4. The Main Automation Loop
# ---------------------------------------------------------
def main():
    output_filename = 'leads_output.xlsx'
    input_filename = 'search_inputs.xlsx'
    
    print(f"Reading inputs from '{input_filename}'...")
    try:
        inputs_df = pd.read_excel(input_filename)
    except FileNotFoundError:
        print(f"Error: Could not find '{input_filename}'. Please create it first.")
        return
        
    existing_searches = set()
    existing_leads = []
    
    if os.path.exists(output_filename):
        print(f"Found existing '{output_filename}'. Loading older entries...")
        try:
            existing_df = pd.read_excel(output_filename)
            existing_leads = existing_df.to_dict('records')
            
            for lead in existing_leads:
                search_id = (
                    str(lead.get('Search Sector', '')).strip().lower(),
                    str(lead.get('Search Target Company', '')).strip().lower(),
                    str(lead.get('Search Target Title', '')).strip().lower(),
                    str(lead.get('Search City', '')).strip().lower()
                )
                existing_searches.add(search_id)
            print(f"Loaded {len(existing_leads)} older leads from previous runs.")
        except Exception as e:
            print(f"Warning: Could not read existing output file ({e}). Starting fresh.")

    new_leads_counter = 0

    for index, row in inputs_df.iterrows():
        current_search_id = (
            str(row['Sector']).strip().lower(),
            str(row['Company']).strip().lower(),
            str(row['Job Title']).strip().lower(),
            str(row['City']).strip().lower()
        )
        
        if current_search_id in existing_searches:
            print(f"Skipping: {row['Job Title']} at {row['Company']} in {row['City']} (Already Processed)")
            continue
            
        print(f"\nProcessing NEW Row: {row['Job Title']} at {row['Company']} in {row['City']} (Page {row['Search Page']})")
        
        profiles = get_linkedin_profiles(
            row['Sector'], 
            row['Company'], 
            row['Job Title'], 
            row['City'], 
            row['Search Page']
        )
        
        for profile in profiles:
            title = profile.get("title", "")
            snippet = profile.get("snippet", "")
            link = profile.get("link", "")
            
            print(f"  -> Processing Profile: {title[:30]}...")
            
            try:
                print("  -> Waiting 15 seconds for API rate limit...")
                time.sleep(15)
                
                ai_response = generate_lead_data(title, snippet)
                
                parts = ai_response.split("|||")
                
                if len(parts) >= 4:
                    lead_name = parts[0].strip()
                    lead_company = parts[1].strip()
                    lead_position = parts[2].strip()
                    pitch = parts[3].strip()
                else:
                    lead_name = "Data Error"
                    lead_company = "Data Error"
                    lead_position = "Data Error"
                    pitch = ai_response.strip()
                
                # FIX: Pass the raw URL string. xlsxwriter will automatically make it clickable.
                valid_link = link if link else "No Link Found"

                # REORDERED COLUMNS: Company First
                existing_leads.append({
                    "Company Name": lead_company,
                    "Name": lead_name,
                    "Designation": lead_position,
                    "LinkedIn Link": valid_link, 
                    "Message": pitch,
                    # --- Search data kept at the end for script memory ---
                    "Search Sector": row['Sector'],
                    "Search Target Company": row['Company'],
                    "Search Target Title": row['Job Title'],
                    "Search City": row['City']
                })
                new_leads_counter += 1
                
            except Exception as e:
                print(f"  -> Error processing AI data for {title}: {e}")
                
    if new_leads_counter > 0:
        print(f"\nSaving master sheet with {new_leads_counter} new entries added...")
        output_df = pd.DataFrame(existing_leads)
        
        
        try:
            # FIX 3: Use xlsxwriter engine to format columns automatically
            with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
                output_df.to_excel(writer, index=False, sheet_name='Leads')
                
                workbook = writer.book
                worksheet = writer.sheets['Leads']
                
                wrap_format = workbook.add_format({'text_wrap': True, 'valign': 'top'})
                std_format = workbook.add_format({'valign': 'top'})
                
                # Updated widths for the new column order
                worksheet.set_column(0, 0, 20, std_format)   # Col A: Company Name
                worksheet.set_column(1, 1, 15, std_format)   # Col B: Name
                worksheet.set_column(2, 2, 20, std_format)   # Col C: Designation
                worksheet.set_column(3, 3, 15, std_format)   # Col D: LinkedIn Link
                worksheet.set_column(4, 4, 60, wrap_format)  # Col E: Message (Wide & Wrapped)
                worksheet.set_column(5, 8, 15, std_format)   # Col F-I: Background search data
                
            print(f"Done! '{output_filename}' updated successfully.")
        except PermissionError:
            print(f"\nCRITICAL ERROR: Could not save data! Please CLOSE '{output_filename}' if it is open in Excel and try running the script again.")
    else:
        print("\nNo new entries found in the input sheet to process.")

if __name__ == "__main__":
    main()