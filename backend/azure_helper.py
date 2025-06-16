import requests
from datetime import datetime, timedelta

# Get Azure Token
def get_azure_token(tenant_id, client_id, client_secret):
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    payload = {
        "client_id": client_id,
        "client_secret": client_secret,
        "grant_type": "client_credentials",
        "scope": "https://management.azure.com/.default"
    }
    response = requests.post(url, data=payload)
    response.raise_for_status()
    return response.json()["access_token"]

# Fetch Azure Resources
def fetch_resources(subscription_id, tenant_id, client_id, client_secret):
    token = get_azure_token(tenant_id, client_id, client_secret)
    url = f"https://management.azure.com/subscriptions/{subscription_id}/resources?api-version=2021-04-01"
    headers = {"Authorization": f"Bearer {token}"}

    resources = []
    while url:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        data = response.json()

        for item in data.get("value", []):
            resources.append({
                "subscription_id": subscription_id,
                "resource_group": item.get("resourceGroup"),
                "name": item.get("name"),
                "type": item.get("type"),
                "location": item.get("location"),
                "id": item.get("id")
            })

        url = data.get("nextLink")

    return resources

# Fetch Azure Cost for given resources
def fetch_cost(subscription_id, tenant_id, client_id, client_secret, resource_id):
    token = get_azure_token(tenant_id, client_id, client_secret)
    url = f"https://management.azure.com/subscriptions/{subscription_id}/providers/Microsoft.CostManagement/query?api-version=2023-03-01"

    end_date = datetime.utcnow().date()
    start_date = end_date - timedelta(days=30)
    body = {
        "type": "Usage",
        "timeframe": "Custom",
        "timePeriod": {
            "from": str(start_date),
            "to": str(end_date)
        },
        "dataset": {
            "granularity": "None",
            "aggregation": {
                "totalCost": {
                    "name": "PreTaxCost",
                    "function": "Sum"
                }
            },
            "filter": {
                "dimensions": {
                    "name": "ResourceId",
                    "operator": "In",
                    "values": [resource_id]
                }
            }
        }
    }

    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    response = requests.post(url, json=body, headers=headers)
    response.raise_for_status()

    data = response.json()
    rows = data.get("properties", {}).get("rows", [])

    if rows:
        return float(rows[0][0])
    else:
        return 0.0
