# Usage Guide

## Workflow Overview

```
0. Authenticate (First Run) → 1. Define Agents (CSV) → 2. Provision Users → 3. Assign Licenses → 4. Export Config
```

## Step 0: Authentication (First Run)

### Device Code Authentication

The first time you run the provisioning tool, you'll need to authenticate:

```bash
npm run provision
```

You'll see authentication prompts:

```
🔐 Microsoft 365 Authentication Required

This application requires admin permissions to provision users.
You will be prompted to sign in with your Microsoft 365 admin account.

📋 To sign in, use a web browser to open the page:
   https://microsoft.com/devicelogin

📝 And enter the code: A1B2C3D4

⏳ Waiting for authentication...
   (This window will update automatically when you complete sign-in)
```

**Steps:**
1. Open the URL in your browser: `https://microsoft.com/devicelogin`
2. Enter the code displayed (e.g., `A1B2C3D4`)
3. Sign in with your Microsoft 365 admin account
4. Grant consent to the requested permissions
5. Return to the CLI - it will continue automatically

**Authentication complete:**
```
✅ Authentication successful!

   Signed in as: admin@yourdomain.onmicrosoft.com
   Token expires: 1/21/2026, 11:30:00 AM

✅ Using cached authentication token
```

### Token Caching

After first authentication:
- Token is cached in `~/.m365-provision/token-cache.json`
- Future runs use the cached token (no re-authentication needed)
- Token automatically refreshes when expired (up to 90 days)

### Authentication Commands

```bash
# Force re-authentication (ignore cache)
npm run provision -- --auth

# Clear cached tokens and logout
npm run provision -- --logout

# Check cached token status
node dist/auth/token-cache.js info
```

## Step 1: Define Agents

### Create Agent CSV

Edit `config/agents-template.csv` with your desired agent definitions:

**Standard CSV (core attributes):**
```csv
name,email,role,department
Sarah Chen,sarah.chen@yourdomain.com,CEO,Executive
Michael Rodriguez,michael.rodriguez@yourdomain.com,CTO,Engineering
```

**Extended CSV (beta endpoints):**
```csv
name,email,role,department,employeeType,companyName,officeLocation
Sarah Chen,sarah.chen@yourdomain.com,CEO,Executive,Employee,Contoso Ltd,Building 1
Michael Rodriguez,michael.rodriguez@yourdomain.com,CTO,Engineering,Employee,Contoso Ltd,Building 2
Robert Taylor,robert.taylor@yourdomain.com,QA Engineer,Engineering,Contractor,Fabrikam Inc,Remote
```

### CSV Format

#### Required Columns

| Column     | Description                           | Example                  |
|------------|---------------------------------------|--------------------------|
| name       | Full name of the agent                | Sarah Chen               |
| email      | Email address (use your domain)       | sarah@domain.com         |
| role       | Job title/role                        | CEO                      |
| department | Department/team                       | Executive                |

#### Optional Columns (Beta endpoints - always enabled)

| Column         | Description                        | Example           |
|----------------|------------------------------------|--------------------|
| employeeType   | Type of employee                   | Employee, Contractor, Intern |
| companyName    | Organization or company name       | Contoso Ltd        |
| officeLocation | Physical office location           | Building 1, Remote |

**Note**: Beta endpoints are always used. Extended columns are supported by default.

### Agent Role Guidelines

**Executive Roles** (1-2 agents):
- CEO, COO, CFO
- High-level decision makers
- Frequent meeting creators

**Engineering Roles** (5-8 agents):
- Developers (Frontend, Backend, Full-stack)
- DevOps Engineers
- QA Engineers
- Frequent email communicators

**Product Roles** (2-3 agents):
- Product Manager
- Product Owner
- UX/UI Designer

**Business Roles** (2-4 agents):
- Marketing Manager
- Sales Director
- Customer Success Manager

**Support Roles** (1-2 agents):
- HR Manager
- Office Manager
- IT Support

## Step 2: Run Provisioning

### Basic Usage

```bash
npm run provision
```

This will:
1. Read `config/agents-template.csv`
2. Create Microsoft 365 user accounts
3. Assign licenses
4. Export to `output/agents-config.json`

### Dry Run (Preview Only)

```bash
npm run provision -- --dry-run
```

Previews what will be created without making changes.

### Options

```bash
npm run provision -- [options]

Options:
  --dry-run              Preview changes without creating users
  --csv <path>           Use custom CSV file (default: config/agents-template.csv)
  --output <path>        Custom output path (default: output/agents-config.json)
  --skip-licenses        Skip license assignment (users only)
  --auth                 Force re-authentication (ignore cached token)
  --logout               Clear cached authentication token
  --force                Overwrite existing users (dangerous!)
  --help, -h             Show help message
```

### Using Beta Features (Extended Attributes)

To provision users with extended attributes (employeeType, companyName, officeLocation):

1. Add beta columns to your CSV:
   ```csv
   name,email,role,department,employeeType,companyName,officeLocation
   Sarah Chen,sarah.chen@domain.com,CEO,Executive,Employee,Contoso Ltd,Building 1
   ```

2. Run provisioning:
   ```bash
   npm run provision
   ```

3. Output will indicate beta usage:
   ```
   Configuration:
     CSV: config/agents-template.csv
     Beta Features: ✓ Enabled

   ✓ Created user [beta]: Sarah Chen (sarah.chen@domain.com)
   ```

**Important Notes:**
- Beta endpoints provide preview features but may change without notice
- This project does **not** fall back to v1.0
- See [docs/BETA-API-GUIDE.md](./docs/BETA-API-GUIDE.md) for detailed beta documentation

### Example Commands

```bash
# Provisioning (beta endpoints)
npm run provision

# Dry run to preview changes
npm run provision -- --dry-run

# Use custom CSV file
npm run provision -- --csv config/custom-agents.csv

# Skip license assignment (users only)
npm run provision -- --skip-licenses

# Force re-authentication
npm run provision -- --auth
```

### Examples

```bash
# Use custom CSV file
npm run provision -- --csv config/my-agents.csv

# Preview what will be created
npm run provision -- --dry-run

# Create users without license assignment (testing)
npm run provision -- --skip-licenses
```

## Step 3: Verify Provisioning

### Check Created Users

```bash
npm run list-users
```

Output:
```
✓ Found 12 provisioned agent users:

Name                    Email                              License    Created
Sarah Chen              sarah.chen@domain.com              E3         2024-01-15
Michael Rodriguez       michael.rodriguez@domain.com       E3         2024-01-15
...
```

### Test User Access

Users can sign in at https://office.com with auto-generated passwords (saved in output).

### Verify Mailboxes

Mailbox provisioning can take 15-30 minutes. Check status:

```bash
npm run check-mailboxes
```

## Step 4: Review Output

### Output Format

The tool exports user configuration to `output/agents-config.json`:

```json
{
  "agents": [
    {
      "name": "Sarah Chen",
      "email": "sarah.chen@yourdomain.com",
      "role": "CEO",
      "department": "Executive",
      "userId": "a1b2c3d4-...",
      "password": "AutoGen123!@#",
      "employeeType": "Employee",
      "companyName": "Contoso Ltd",
      "officeLocation": "Building 1",
      "createdAt": "2024-01-15T10:30:00Z"
    }
  ],
  "summary": {
    "totalAgents": 12,
    "successfulProvisions": 12,
    "failedProvisions": 0,
    "generatedAt": "2024-01-15T10:30:00Z"
  }
}
```

### Using the Output

The generated `output/agents-config.json` file contains all user details and can be used by other applications.

Additionally, user passwords are saved to `output/passwords.txt` for reference.

**⚠️ Security Note**: Keep the passwords file secure and never commit it to git!

## Common Workflows

### Scenario 1: First-Time Setup (10 Agents)

```bash
# 1. Edit CSV with 10 agents
nano config/agents-template.csv

# 2. Preview what will be created
npm run provision -- --dry-run

# 3. Create users and assign licenses
npm run provision

# 4. Verify
npm run list-users

# 5. Review output
cat output/agents-config.json
```

### Scenario 2: Add More Agents Later

```bash
# 1. Add new rows to CSV
nano config/agents-template.csv

# 2. Provision only new users (existing users skipped automatically)
npm run provision

# 3. Review updated configuration
cat output/agents-config.json
```

### Scenario 3: Clean Up Test Users

```bash
# Remove all provisioned agent users
npm run cleanup

# WARNING: This is destructive! Users and mailboxes will be deleted.
```

## Troubleshooting

### Issue: "User already exists"

**Solution**: Provisioning script skips existing users by default. To force recreation, use `--force` flag (not recommended).

### Issue: "License assignment failed"

**Possible causes**:
- No licenses available
- User already has a license
- Wrong LICENSE_SKU_ID

**Solution**:
```bash
# Check available licenses
npm run list-licenses

# Update licenses for all users
npm run update-licenses -- --csv config/your-users.csv
```

### Issue: "Need to add a new license to existing users"

When you add a new license to `LICENSE_SKU_IDS` in `.env` (e.g., adding Copilot license), existing users won't automatically get it. The provisioning command only assigns licenses during user creation.

**Solution**: Use the dedicated license update command:
```bash
# Preview which licenses would be added
npm run update-licenses -- --dry-run --csv config/your-users.csv

# Apply the license updates
npm run update-licenses -- --csv config/your-users.csv
```

This command:
- Reads users from your CSV
- Checks their current licenses
- Adds any missing licenses from `LICENSE_SKU_IDS`
- Skips licenses already assigned (idempotent)

### Issue: "CSV parsing error"

**Causes**:
- Invalid CSV format
- Missing required columns
- Special characters in names

**Solution**:
1. Verify CSV has header row: `name,email,role,department`
2. Ensure no empty rows
3. Use UTF-8 encoding
4. Quote fields with commas: `"Last, First"`

## Best Practices

### 1. Start Small
- Provision 3-5 test agents first
- Verify everything works
- Scale up to 10-20 agents

### 2. Use Descriptive Names
- Use realistic names for agents
- Avoid generic names like "Test User 1"
- Helps with realistic simulations

### 3. Organize by Department
- Group agents by department in CSV
- Makes it easier to create team scenarios
- CEO → Managers → Individual Contributors

### 4. Backup Configuration
```bash
# Backup generated config before regenerating
cp output/agents-config.json output/agents-config.backup.json
```

### 5. Version Control CSV
- Commit `agents-template.csv` to git
- Don't commit `output/agents-config.json` (contains passwords)
- Add `output/` to `.gitignore`

## Next Steps

Once users are provisioned:
1. Review `output/agents-config.json` for user details
2. Securely store `output/passwords.txt`
3. Users can sign in to Microsoft 365 at https://office.com
4. Applications can consume the JSON configuration as needed
5. Monitor user activity in Microsoft 365 admin center
