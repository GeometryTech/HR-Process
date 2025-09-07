import os
import re
import pandas as pd
from PyPDF2 import PdfReader
from collections import Counter
import spacy
import matplotlib.pyplot as plt

nlp = spacy.load("en_core_web_sm")

# Technology keywords
TECH_KEYWORDS = [
    # Programming & scripting
    "python", "java", "c++", "c#", "javascript", "typescript", "powershell", "bash", "shell", "sql", "nosql",
    "mongodb", "mysql", "postgresql", "go", "ruby", "perl", "php", "swift", "objective-c", "scala", "r",
    # Cloud platforms
    "aws", "amazon web services", "azure", "microsoft azure", "gcp", "google cloud", "cloud", "cloud computing",
    # Microsoft technologies
    "office 365", "microsoft 365", "exchange", "sharepoint", "onedrive", "teams", "intune", "active directory",
    "windows server", "windows 10", "windows 11", "microsoft defender", "microsoft endpoint manager",
    "microsoft sql server", "outlook", "visio", "power bi", "powerapps", "microsoft dynamics", "microsoft edge",
    "system center", "sccm", "azure ad", "azure devops", "azure functions", "azure logic apps", "azure storage",
    "azure virtual machines", "azure networking", "azure backup", "azure site recovery", "azure monitor",
    "azure security center", "azure policy", "azure automation", "azure app service", "azure kubernetes service",
    # Virtualization & infrastructure
    "vmware", "hyper-v", "virtualization", "citrix", "remote desktop", "rdp", "vpn", "firewall", "networking",
    "dns", "dhcp", "tcp/ip", "lan", "wan", "switch", "router", "load balancer", "network printer", "firmware",
    # Security
    "security", "cybersecurity", "endpoint protection", "antivirus", "mdr", "siem", "sentinelone", "huntress",
    "threatlocker", "socs 2", "soc 2", "compliance", "encryption", "multi-factor authentication", "mfa",
    "email security", "dns filter", "vulnerability scanner", "av", "vpn", "ssl", "tls", "zero trust",
    # DevOps & automation
    "devops", "jenkins", "terraform", "ansible", "docker", "kubernetes", "automation", "scripting",
    "ci/cd", "gitlab", "github actions", "octopus deploy", "configuration management",
    # Data & analytics
    "data science", "machine learning", "deep learning", "pandas", "numpy", "scikit-learn", "tensorflow",
    "pytorch", "spark", "hadoop", "tableau", "reporting", "analytics", "etl", "data warehouse", "bigquery",
    "databricks", "power bi", "business intelligence",
    # Web & frontend
    "react", "node.js", "angular", "html", "css", "sass", "graphql", "rest", "api", "bootstrap", "vue.js",
    # MSP tools & platforms
    "rmm", "psa", "connectwise", "autotask", "datto", "pax8", "halo", "missive", "ticketing", "remote monitoring",
    "backup", "disaster recovery", "business continuity", "sla", "service desk", "help desk", "solarwinds",
    "manage engine", "ninjaone", "kaseya", "logicmonitor", "auvik", "itglue", "documentation", "monitoring",
    # Other
    "linux", "macos", "apple", "android", "ios", "virtual machine", "network printer", "firmware", "dns filter",
    "salesforce", "zendesk", "servicenow", "jira", "confluence", "slack", "teams", "zoom", "webex"
]

def extract_text_from_pdf(pdf_path):
    reader = PdfReader(pdf_path)
    text = ""
    for page in reader.pages:
        text += page.extract_text() or ""
    return text

def extract_contact_info(text):
    email_match = re.search(r'[\w\.-]+@[\w\.-]+', text)
    email = email_match.group(0) if email_match else ""
    phone_match = re.search(r'(\+?\d{1,2}[\s-]?)?(\(?\d{3}\)?[\s-]?)?\d{3}[\s-]?\d{4}', text)
    phone = phone_match.group(0) if phone_match else ""
    doc = nlp(text)
    name = ""
    for ent in doc.ents:
        if ent.label_ == "PERSON":
            name = ent.text
            break
    return {"Name": name, "Email": email, "Phone": phone}

def extract_technology_skills(text):
    text_lower = text.lower()
    skills = set()
    for keyword in TECH_KEYWORDS:
        if keyword in text_lower:
            skills.add(keyword.title())
    return sorted(skills)

def is_safe_path(basedir, path):
    return os.path.realpath(path).startswith(os.path.realpath(basedir))

def save_excel_with_totals(df, excel_file):
    # Flatten skills and remove duplicates
    all_skills = [skill.strip() for s in df["Skills"] for skill in str(s).split(",") if skill.strip()]
    unique_skills = sorted(list(set(all_skills)))

    # Build skill matrix
    df_matrix = pd.DataFrame(0, index=df["Name"], columns=unique_skills)

    # Fill matrix
    for i, row in df.iterrows():
        for skill in str(row["Skills"]).split(","):
            skill = skill.strip()
            if skill:
                df_matrix.at[row["Name"], skill] += 1

    # Add totals
    df_matrix["TOTAL"] = df_matrix.sum(axis=1)
    df_matrix.loc["TOTAL"] = df_matrix.sum()
    df_matrix.to_excel(excel_file)
    return df_matrix

def plot_skill_distribution(skill_counts, output_file="skill_distribution.png"):
    skills, counts = zip(*skill_counts.most_common(20))
    import matplotlib.pyplot as plt
    plt.figure(figsize=(12, 6))
    plt.bar(skills, counts, color="skyblue")
    plt.xticks(rotation=45, ha="right")
    plt.ylabel("Frequency")
    plt.title("Top 20 Technology Skills")
    plt.tight_layout()
    plt.savefig(output_file)
    plt.close()

def main(folder_path):
    folder_path = os.path.abspath(folder_path)
    if not os.path.isdir(folder_path):
        print("Invalid folder path.")
        return

    data = []
    all_skills = []

    for filename in os.listdir(folder_path):
        if filename.lower().endswith(".pdf"):
            pdf_path = os.path.join(folder_path, filename)
            if not is_safe_path(folder_path, pdf_path):
                print(f"Skipping unsafe file path: {pdf_path}")
                continue
            try:
                text = extract_text_from_pdf(pdf_path)
                contact = extract_contact_info(text)
                skills = extract_technology_skills(text)
                all_skills.extend(skills)
                data.append({
                    "Candidate": filename,
                    "Name": contact["Name"],
                    "Email": contact["Email"],
                    "Phone": contact["Phone"],
                    "Skills": ", ".join(skills)
                })
            except Exception as e:
                print(f"Error reading {filename}: {e}")

    # Save CSV
    df = pd.DataFrame(data)
    df.to_csv("candidate_skills.csv", index=False)
    print("Spreadsheet saved as candidate_skills.csv")

    # Save Excel with totals
    excel_file = "candidate_skill_matrix.xlsx"
    save_excel_with_totals(df, excel_file)
    print(f"Excel skill matrix saved as {excel_file}")

    # Plot top skills
    from collections import Counter
    skill_counts = Counter(all_skills)
    plot_skill_distribution(skill_counts)
    print("Skill distribution chart saved as skill_distribution.png")

    # Print top 20 skills
    print("\nTechnology Skill Distribution:")
    for skill, count in skill_counts.most_common(20):
        print(f"{skill}: {count}")

if __name__ == "__main__":
    folder = input("Enter the path to the folder of resumes: ")
    main(folder)
