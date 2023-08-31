import requests
from openpyxl import Workbook
from openpyxl.styles import Font

API_URL = "https://api.github.com"

def write_secret_scanning_alerts_data_to_excel(repo_owner: str, repo_name: str, token: str, output_file: str):
    url = f"{API_URL}/repos/{repo_owner}/{repo_name}/secret-scanning/alerts"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json"
    }

    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        try:
            secret_alerts = response.json()

            if isinstance(secret_alerts, list):
                if secret_alerts:
                    workbook = Workbook()
                    sheet = workbook.active
                    sheet.append(["number", "created_at", "updated_at", "url", "html_url", "locations_url", "state", 'secret_type',
                                 "secret_type_display_name", "secret", "resolution", "resolved_at", "resolution_comment",
                                 "push_protection_bypassed", "push_protection_bypassed_by", "push_protection_bypassed_at",
                                 "resolved_by_login", "resolved_by_id", "resolved_by_node_id", "resolved_by_avatar_url",
                                 "resolved_by_gravatar_id", "resolved_by_url", "resolved_by_html_url", "resolved_by_followers_url",
                                 "resolved_by_following_url", "resolved_by_gists_url", "resolved_by_starred_url",
                                 "resolved_by_subscriptions_url", "resolved_by_organizations_url", "resolved_by_repos_url",
                                 "resolved_by_events_url", "resolved_by_received_events_url", "resolved_by_type", "resolved_by_site_admin"])
                    
                    for alert in secret_alerts:
                        resolved_by = alert.get("resolved_by")
                        resolved_by_login = resolved_by.get("login") if resolved_by else None
                        resolved_by_id = resolved_by.get("id") if resolved_by else None
                        resolved_by_node_id = resolved_by.get("node_id") if resolved_by else None
                        resolved_by_avatar_url = resolved_by.get("avatar_url") if resolved_by else None
                        resolved_by_gravatar_id = resolved_by.get("gravatar_id") if resolved_by else None
                        resolved_by_url = resolved_by.get("url") if resolved_by else None
                        resolved_by_html_url = resolved_by.get("html_url") if resolved_by else None
                        resolved_by_followers_url = resolved_by.get("followers_url") if resolved_by else None
                        resolved_by_following_url = resolved_by.get("following_url") if resolved_by else None
                        resolved_by_gists_url = resolved_by.get("gists_url") if resolved_by else None
                        resolved_by_starred_url = resolved_by.get("starred_url") if resolved_by else None
                        resolved_by_subscriptions_url = resolved_by.get("subscriptions_url") if resolved_by else None
                        resolved_by_organizations_url = resolved_by.get("organizations_url") if resolved_by else None
                        resolved_by_repos_url = resolved_by.get("repos_url") if resolved_by else None
                        resolved_by_events_url = resolved_by.get("events_url") if resolved_by else None
                        resolved_by_received_events_url = resolved_by.get("received_events_url") if resolved_by else None
                        resolved_by_type = resolved_by.get("type") if resolved_by else None
                        resolved_by_site_admin = resolved_by.get("site_admin") if resolved_by else None
                        
                        sheet.append([
                            alert.get("number"),
                            alert.get("created_at"),
                            alert.get("updated_at"),
                            alert.get("url"),
                            alert.get("html_url"),
                            alert.get("locations_url"),
                            alert.get("state"),
                            alert.get("secret_type"),
                            alert.get("secret_type_display_name"),
                            alert.get("secret"),
                            alert.get("resolution"),
                            alert.get("resolved_at"),
                            alert.get("resolution_comment"),
                            alert.get("push_protection_bypassed"),
                            alert.get("push_protection_bypassed_by"),
                            alert.get("push_protection_bypassed_at"),
                            resolved_by_login,
                            resolved_by_id,
                            resolved_by_node_id,
                            resolved_by_avatar_url,
                            resolved_by_gravatar_id,
                            resolved_by_url,
                            resolved_by_html_url,
                            resolved_by_followers_url,
                            resolved_by_following_url,
                            resolved_by_gists_url,
                            resolved_by_starred_url,
                            resolved_by_subscriptions_url,
                            resolved_by_organizations_url,
                            resolved_by_repos_url,
                            resolved_by_events_url,
                            resolved_by_received_events_url,
                            resolved_by_type,
                            resolved_by_site_admin
                        ])
                        
                    # Apply formatting to header row
                    header_row = sheet[1]
                    for cell in header_row:
                        cell.font = Font(bold=True)

                    workbook.save(output_file)
                    print(f"Successfully wrote secret scanning alerts")
                else:
                    print(f"Repository has no secret scanning alerts.")
            else:
                print(f"Invalid data received from the API. Expected a list of alerts.")
        except Exception as e:
            print(f"An error occurred while processing data: {e}")
    elif response.status_code == 404:
        print(f"Repository or its secret scanning alerts metadata does not exist.")
    else:
        print(f"Failed to get secret scanning alerts metadata. Status code: {response.status_code}")

# Example usage
repo_owner = "DevanshiAgnihotri"
repo_name = "ghas_fp"
token = "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"
output_file = "C:/Users/dagnihotri/Documents/KPMG/mmmyyy_output_file.xlsx"
azure_ad_key1 = "p9m-kN6B-l5D7oZ4qR0n-A1tV8c3X2fU"
token = 'ghp_b7f776e7fa5d4fddb50bf69cfad8d846e'

try:
    write_secret_scanning_alerts_data_to_excel(repo_owner, repo_name, token, output_file)
except Exception as e:
    print(f"error: {e}")

