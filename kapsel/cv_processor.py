import json
import pandas as pd


def process_cv_json(json_response_str: str) -> dict[str, pd.DataFrame] | None:
    """Parse the structured JSON produced by the CV-extraction prompt and return

    a dict of DataFrames ready for downstream HR-system consumption.

    Keys returned:
        "summary"     – one-row candidate overview
        "education"   – one row per education entry
        "experience"  – one row per work-experience entry
        "skills"      – one row listing hard skills, soft skills, certifications
    """
    try:
        data = json.loads(json_response_str)
    except json.JSONDecodeError as e:
        print(f"Gagal memparsing JSON: {e}")
        return None

    personal = data.get("personal_info", {})

    # 1. Candidate summary row
    summary_row = {
        "Full Name": personal.get("full_name", "N/A"),
        "Email": personal.get("email", "N/A"),
        "Phone Number": personal.get("phone_number", "N/A"),
        "Location": personal.get("location", "N/A"),
        "LinkedIn": personal.get("linkedin_url", "N/A"),
        "Summary": data.get("summary", "N/A"),
    }
    df_summary = pd.DataFrame([summary_row])

    # 2. Education entries
    education_rows = []
    for edu in data.get("education", []):
        education_rows.append({
            "Institution": edu.get("institution", "N/A"),
            "Degree": edu.get("degree", "N/A"),
            "Field of Study": edu.get("field_of_study", "N/A"),
            "Start Year": edu.get("start_year", "N/A"),
            "End Year": edu.get("end_year", "N/A"),
        })
    df_education = (
        pd.DataFrame(education_rows)
        if education_rows
        else pd.DataFrame(
            columns=[
                "Institution",
                "Degree",
                "Field of Study",
                "Start Year",
                "End Year",
            ]
        )
    )

    # 3. Work experience entries
    experience_rows = []
    for exp in data.get("work_experience", []):
        responsibilities = exp.get("key_responsibilities", [])
        experience_rows.append({
            "Company": exp.get("company", "N/A"),
            "Role": exp.get("role", "N/A"),
            "Start Date": exp.get("start_date", "N/A"),
            "End Date": exp.get("end_date", "N/A"),
            "Key Responsibilities": " | ".join(responsibilities)
            if responsibilities
            else "N/A",
        })
    df_experience = (
        pd.DataFrame(experience_rows)
        if experience_rows
        else pd.DataFrame(
            columns=[
                "Company",
                "Role",
                "Start Date",
                "End Date",
                "Key Responsibilities",
            ]
        )
    )

    # 4. Skills & certifications row
    certifications = data.get("certifications", [])
    certifications = [c for c in certifications if c != "N/A"] or ["N/A"]

    skills_row = {
        "Hard Skills": ", ".join(data.get("hard_skills", [])) or "N/A",
        "Soft Skills": ", ".join(data.get("soft_skills", [])) or "N/A",
        "Certifications": ", ".join(certifications),
    }
    df_skills = pd.DataFrame([skills_row])

    return {
        "summary": df_summary,
        "education": df_education,
        "experience": df_experience,
        "skills": df_skills,
    }


def save_to_excel(
    dataframes: dict[str, pd.DataFrame],
    filepath: str = "extracted_cv_result.xlsx",
) -> None:
    """Write each DataFrame section to a separate sheet in one Excel workbook."""
    with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
        for sheet_name, df in dataframes.items():
            df.to_excel(writer, sheet_name=sheet_name.capitalize(), index=False)
    print(f"File berhasil disimpan ke '{filepath}'")


if __name__ == "__main__":
    sample_json_output = """
    {
      "personal_info": {
        "full_name": "Budi Santoso",
        "email": "budi.santoso@email.com",
        "phone_number": "081234567890",
        "location": "Yogyakarta, Indonesia",
        "linkedin_url": "linkedin.com/in/budisantoso"
      },
      "summary": "Experienced Software Engineer with 4 years in backend API development using Python and Java. Skilled in cloud infrastructure on AWS and collaborative agile delivery. Passionate about building scalable, maintainable systems.",
      "education": [
        {
          "institution": "Universitas Gadjah Mada",
          "degree": "Bachelor of Engineering",
          "field_of_study": "Computer Science",
          "start_year": "2016",
          "end_year": "2020"
        }
      ],
      "work_experience": [
        {
          "company": "PT Teknologi Nusantara",
          "role": "Backend Engineer",
          "start_date": "07/2020",
          "end_date": "Present",
          "key_responsibilities": [
            "Designed and maintained RESTful APIs serving 500k+ daily users",
            "Migrated monolithic services to microservices architecture on AWS",
            "Reduced average API response time by 35% via query optimization"
          ]
        }
      ],
      "hard_skills": ["Python", "Java", "SQL", "AWS", "Docker", "REST API"],
      "soft_skills": ["Problem Solving", "Team Collaboration", "Communication"],
      "certifications": ["AWS Certified Developer – Associate", "N/A"]
    }
    """

    result = process_cv_json(sample_json_output)
    if result is not None:
        for section, df in result.items():
            print(f"\n=== {section.upper()} ===")
            print(df.to_string(index=False))

        save_to_excel(result)