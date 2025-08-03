import boto3
import pandas as pd
from botocore.exceptions import ClientError

def assume_role(account_id, role_name="System_Admin"):
    """
    Assumes a role in the target AWS account and returns temporary credentials.
    """
    try:
        sts_client = boto3.client("sts")
        response = sts_client.assume_role(
            RoleArn=f"arn:aws:iam::{account_id}:role/{role_name}",
            RoleSessionName="PatchBaselineFetcher",
            DurationSeconds=900,
        )
        return response["Credentials"]
    except ClientError as e:
        print(f"Error assuming role in account {account_id}: {e}")
        return None

def get_windows_patch_baselines(session, region):
    """
    Fetches patch baseline details specifically for Windows OS,
    including approval days, using the provided boto3 session and region.
    """
    try:
        ssm_client = session.client("ssm", region_name=region)
        paginator = ssm_client.get_paginator("describe_patch_baselines")
        
        response_iterator = paginator.paginate(
            Filters=[
                {
                    'Key': 'OPERATING_SYSTEM', 
                    'Values': ['WINDOWS','AMAZON_LINUX_2']
                }
            ]
        )
        
        baselines_with_details = []
        for page in response_iterator:
            for baseline_identity in page["BaselineIdentities"]:
                baseline_id = baseline_identity["BaselineId"]
                baseline_name = baseline_identity["BaselineName"]
                
                # Use get_patch_baseline to get more details including ApprovalRules
                baseline_details = ssm_client.get_patch_baseline(BaselineId=baseline_id)
                
                approval_days = []
                if "ApprovalRules" in baseline_details and "PatchRules" in baseline_details["ApprovalRules"]:
                    for rule in baseline_details["ApprovalRules"]["PatchRules"]:
                        if "ApproveAfterDays" in rule:
                            approval_days.append(str(rule["ApproveAfterDays"])) # Convert to string for easier CSV/Excel handling

                baselines_with_details.append({
                    "BaselineId": baseline_id,
                    "BaselineName": baseline_name,
                    "OperatingSystem": baseline_identity.get("OperatingSystem", "N/A"),
                    "DefaultBaseline": baseline_identity.get("DefaultBaseline", False),
                    "ApprovalDays": ", ".join(approval_days) if approval_days else "N/A"
                })
        return baselines_with_details
    except ClientError as e:
        print(f"Error fetching patch baselines in region {region}: {e}")
        return []

def main():
    input_file = "input.xlsx"
    output_file = "windows_patch_baseline_details.xlsx"
    role_name = "System_Admin"  # Replace with the actual role name

    try:
        df = pd.read_excel(input_file)
    except FileNotFoundError:
        print(f"Error: {input_file} not found. Please create the input Excel file.")
        return
    except Exception as e:
        print(f"Error reading {input_file}: {e}")
        return

    all_windows_patch_baselines = []

    for index, row in df.iterrows():
        account_id = str(row["AccountID"])
        region = row["Region"]
        print(f"Processing account: {account_id}, region: {region}")

        credentials = assume_role(account_id, role_name)
        if credentials:
            session = boto3.Session(
                aws_access_key_id=credentials["AccessKeyId"],
                aws_secret_access_key=credentials["SecretAccessKey"],
                aws_session_token=credentials["SessionToken"],
            )
            windows_patch_baselines = get_windows_patch_baselines(session, region)
            for baseline in windows_patch_baselines:
                baseline["AccountID"] = account_id
                baseline["Region"] = region
            all_windows_patch_baselines.extend(windows_patch_baselines)

    if all_windows_patch_baselines:
        output_df = pd.DataFrame(all_windows_patch_baselines)
        output_df.to_excel(output_file, index=False)
        print(f"Successfully fetched Windows patch baseline details to {output_file}")
    else:
        print("No Windows patch baseline details found.")

if __name__ == "__main__":
    main()
