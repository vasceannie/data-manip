import pandas as pd
import re

def parse_supplier_emails(supplier_data):
    """
    Parse the supplier table and extract all unique emails.
    Returns a set of lowercase emails.
    """
    # Read the CSV data with specific parameters to handle potential issues
    df = pd.read_csv(
        supplier_data,
        encoding='utf-8',
        on_bad_lines='skip',  # Skip problematic lines
        dtype=str  # Read all columns as strings
    )
    
    # Find the email column - look for the specific column name
    email_column = [col for col in df.columns if 'EmailI' in col]
    if not email_column:
        raise ValueError("Could not find email column in supplier data")
    email_column = email_column[0]
    
    # Create a set of all unique emails
    all_emails = set()
    
    # Process each row's emails
    for emails in df[email_column].dropna():
        # Split by semicolon if multiple emails exist
        email_list = [e.strip() for e in emails.split(';')] if ';' in emails else [emails]
        for email in email_list:
            cleaned_email = email.strip().lower()
            if cleaned_email and '@' in cleaned_email:  # Basic email validation
                all_emails.add(cleaned_email)
                
    return all_emails

def parse_user_emails(user_data):
    """
    Parse the user table and extract all unique emails.
    Returns a set of lowercase emails.
    """
    # Read the CSV data with specific parameters
    df = pd.read_csv(
        user_data,
        encoding='utf-8',
        on_bad_lines='skip',  # Skip problematic lines
        dtype=str  # Read all columns as strings
    )
    
    # Use the third column for emails
    email_column = df.columns[2]  # Using index 2 for third column
    if not email_column:
        raise ValueError("Could not find email column in user data")
    
    # Process the emails
    all_emails = set()
    for email in df[email_column].dropna():
        cleaned_email = email.strip().lower()
        if cleaned_email and '@' in cleaned_email:  # Basic email validation
            all_emails.add(cleaned_email)
    
    return all_emails

def compare_emails(supplier_file, user_file):
    """
    Compare emails between supplier and user tables.
    Returns a detailed analysis.
    """
    # Parse emails from both sources
    supplier_emails = parse_supplier_emails(supplier_file)
    user_emails = parse_user_emails(user_file)
    
    # Perform analysis
    emails_in_both = supplier_emails.intersection(user_emails)
    only_in_suppliers = supplier_emails.difference(user_emails)
    only_in_users = user_emails.difference(supplier_emails)
    
    # Prepare analysis results
    return {
        'summary': {
            'total_supplier_emails': len(supplier_emails),
            'total_user_emails': len(user_emails),
            'emails_in_both_count': len(emails_in_both),
            'only_in_suppliers_count': len(only_in_suppliers),
            'only_in_users_count': len(only_in_users)
        },
        'details': {
            'emails_in_both': sorted(emails_in_both),
            'only_in_suppliers': sorted(only_in_suppliers),
            'only_in_users': sorted(only_in_users)
        }
    }
    
def generate_report(analysis):
    """
    Generate a formatted report from the analysis results and save as Excel.
    """
    # Create a new Excel writer object
    with pd.ExcelWriter('email_comparison_report.xlsx') as writer:
        # Create summary DataFrame
        summary_df = pd.DataFrame({
            'Metric': [
                'Total supplier emails',
                'Total user emails',
                'Emails present in both tables',
                'Emails only in supplier table',
                'Emails only in user table'
            ],
            'Count': [
                analysis['summary']['total_supplier_emails'],
                analysis['summary']['total_user_emails'],
                analysis['summary']['emails_in_both_count'],
                analysis['summary']['only_in_suppliers_count'],
                analysis['summary']['only_in_users_count']
            ]
        })
        
        # Create detailed DataFrames
        emails_in_both_df = pd.DataFrame(analysis['details']['emails_in_both'], columns=['Email'])
        only_in_suppliers_df = pd.DataFrame(analysis['details']['only_in_suppliers'], columns=['Email'])
        only_in_users_df = pd.DataFrame(analysis['details']['only_in_users'], columns=['Email'])
        
        # Write each DataFrame to a different worksheet
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        emails_in_both_df.to_excel(writer, sheet_name='Emails in Both', index=False)
        only_in_suppliers_df.to_excel(writer, sheet_name='Only in Suppliers', index=False)
        only_in_users_df.to_excel(writer, sheet_name='Only in Users', index=False)

    # Also return the text report for backward compatibility
    return f"""
Email Comparison Analysis Report
==============================

Summary:
--------
Total supplier emails: {analysis['summary']['total_supplier_emails']}
Total user emails: {analysis['summary']['total_user_emails']}
Emails present in both tables: {analysis['summary']['emails_in_both_count']}
Emails only in supplier table: {analysis['summary']['only_in_suppliers_count']}
Emails only in user table: {analysis['summary']['only_in_users_count']}
"""

def main():
    """
    Main function to run the analysis.
    """
    supplier_file = 'ARContacts.csv'  # This is the supplier file with EmailIDs column
    user_file = 'Susers.csv'  # This has emails in the third column
    
    try:
        analysis = compare_emails(supplier_file, user_file)
        report = generate_report(analysis)
        print(report)
        print("\nReport has been saved to 'email_comparison_report.xlsx'")
            
    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()