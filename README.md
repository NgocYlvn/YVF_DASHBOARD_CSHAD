# CS HAD – YVF Adoption Dashboard

## Deploy on Streamlit Community Cloud
1. Upload all files and folders in this project to one GitHub repository.
2. In Streamlit Community Cloud, choose the repository.
3. Set **Main file path** to `app.py`.
4. Click **Deploy**.

## Update monthly data
Replace the Excel file below with the updated file while keeping the same file name and sheet/header structure:

`data/YVF_Adoption_Dashboard_Source.xlsx`

The dashboard reads these sheets:
- Dashboard_Overview
- Customer_Volume
- Booking_Records
- Onboarded_Customers
- Improvement Proposals
- Customer_Feedback
- User Issues

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```
