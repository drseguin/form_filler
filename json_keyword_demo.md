# JSON Keyword Demonstration

## Sum Example
The sum of monthly totals: {{JSON!sales.json!$.monthly_totals!SUM}}

## Join Example
Team members: {{JSON!users.json!$.names!JOIN(, )}}

## Boolean Transformation Example
System status: {{JSON!status.json!$.system_active!BOOL(Online/Offline)}}

## Component Status Examples
Database status: {{JSON!status.json!$.components.database.active!BOOL(Online/Offline)}}
API status: {{JSON!status.json!$.components.api.active!BOOL(Online/Offline)}}
Frontend status: {{JSON!status.json!$.components.frontend.active!BOOL(Online/Offline)}}

## Additional Examples
Total active users: {{JSON!users.json!$.active_users}}
System maintenance mode: {{JSON!status.json!$.maintenance_mode!BOOL(Yes/No)}}
Last status check: {{JSON!status.json!$.last_status_check}} 