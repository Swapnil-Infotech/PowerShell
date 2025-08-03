import boto3
import pandas as pd
import time
import json  # Import json for parsing PowerShell output

# --- Configuration ---
INSTANCES_FILE = 'instance.xlsx'  # Name of your Excel file with AccountID, Region, InstanceID
OUTPUT_FILE = 'ssm_command_output_structured.xlsx'  # Name of the output Excel file
# PowerShell script to execute (now outputs structured data)
POWERSHELL_SCRIPT = '''
# Example 1: Get service status
$service = Get-Service -Name "WinRM"

# Example 2: Get basic system info
$systemInfo = Get-ComputerInfo -Property OsName, OsVersion, CsProcessors

# Create a custom object with structured output
$output = [PSCustomObject]@{
    "ServiceStatus" = $service.Status
    "OS_Name" = $systemInfo.OsName
    "OS_Version" = $systemInfo.OsVersion
    "CPU_Count" = $systemInfo.CsProcessors
}

# Convert the object to JSON for easy parsing in Python
$output | ConvertTo-Json
'''
# Name of the SSM Automation Execution Role in the *management account* (the account running this script)
# This role needs permissions to assume roles in target accounts and execute SSM commands
AUTOMATION_ASSUME_ROLE_NAME = 'System_Admin'
# Name of the SSM Automation Execution Role in the *target accounts*
# This role needs permissions to execute SSM commands on EC2 instances
EXECUTION_ASSUME_ROLE_NAME = 'System_Admin'


# --- Functions ---

def run_powershell_on_ec2(account_id, region, instance_id, script_content, assumed_role_arn):
    """
    Executes a PowerShell script on a given EC2 instance using SSM Run Command.
    Assumes the specified role to perform actions in the target account.
    Returns the structured output of the script or an error dictionary.
    """
    sts_client = boto3.client('sts')
    try:
        assumed_role_object = sts_client.assume_role(
            RoleArn=assumed_role_arn,
            RoleSessionName="PowerShellRunner"
        )
        credentials = assumed_role_object['Credentials']

        ssm_client = boto3.client(
            'ssm',
            region_name=region,
            aws_access_key_id=credentials['AccessKeyId'],
            aws_secret_access_key=credentials['SecretAccessKey'],
            aws_session_token=credentials['SessionToken']
        )

        print(f"Running PowerShell script on instance {instance_id} in account {account_id}, region {region}...")
        response = ssm_client.send_command(
            InstanceIds=[instance_id],
            DocumentName="AWS-RunPowerShellScript",
            Parameters={'commands': [script_content]}
        )
        command_id = response['Command']['CommandId']

        # Wait for the command to complete and retrieve output
        status = "Pending"
        output_content = ""
        while status not in ["Success", "Failed", "Cancelled", "TimedOut"]:
            time.sleep(5)  # Wait for 5 seconds before checking status again
            command_invocation = ssm_client.get_command_invocation(
                CommandId=command_id,
                InstanceId=instance_id
            )
            status = command_invocation['Status']
            print(f"Command status for {instance_id}: {status}")

        if status == "Success":
            output_content = command_invocation['StandardOutputContent']
            try:
                # Parse the JSON output from PowerShell
                parsed_output = json.loads(output_content)
                return parsed_output
            except json.JSONDecodeError as e:
                error_message = f"Error parsing JSON output from instance {instance_id}: {e}\nRaw output:\n{output_content}"
                print(error_message)
                return {"Error": error_message}
        else:
            error_message = f"SSM command failed or timed out for instance {instance_id}.\nError: {command_invocation.get('StandardErrorContent', 'No error content available')}"
            print(error_message)
            return {"Error": error_message}

    except Exception as e:
        error_message = f"Error running PowerShell script on instance {instance_id} in account {account_id}, region {region}: {e}"
        print(error_message)
        return {"Error": error_message}


# --- Main script execution ---

if __name__ == "__main__":
    print("Starting PowerShell execution across multiple AWS accounts and EC2 instances from Excel file.")

    all_results_df = pd.DataFrame()  # To store all structured command outputs

    try:
        # Read instance details from the Excel file
        df_instances = pd.read_excel(INSTANCES_FILE)

        # Iterate over each row of the DataFrame
        for index, row in df_instances.iterrows():
            account_id = str(row['AccountID'])
            region = row['Region']
            instance_id = row['InstanceID']

            print(f"\nProcessing instance: {instance_id} in account: {account_id}, region: {region}")

            # Construct the ARN for the role to assume in the target account
            assumed_role_arn = f"arn:aws:iam::{account_id}:role/{EXECUTION_ASSUME_ROLE_NAME}"

            structured_output = run_powershell_on_ec2(account_id, region, instance_id, POWERSHELL_SCRIPT,
                                                      assumed_role_arn)

            # Create a dictionary for the current instance's data
            instance_data = {
                'AccountID': account_id,
                'Region': region,
                'InstanceID': instance_id
            }

            # Add the structured output as new columns
            if isinstance(structured_output, dict):
                instance_data.update(structured_output)
            else:
                instance_data['CommandOutput'] = str(structured_output)  # Fallback if not a dict

            # Append as a new row to the DataFrame
            all_results_df = pd.concat([all_results_df, pd.DataFrame([instance_data])], ignore_index=True)


    except FileNotFoundError:
        print(
            f"Error: The Excel file '{INSTANCES_FILE}' was not found. Please ensure it's in the same directory as the script or provide the full path.")
    except Exception as e:
        print(f"An error occurred: {e}")

    # Create a DataFrame from the results and export to Excel
    if not all_results_df.empty:
        try:
            all_results_df.to_excel(OUTPUT_FILE, index=False)
            print(f"\nStructured command outputs successfully written to '{OUTPUT_FILE}'")
        except Exception as e:
            print(f"Error writing output to Excel file: {e}")

    print("\nPowerShell script execution complete.")