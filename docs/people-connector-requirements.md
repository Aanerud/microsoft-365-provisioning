# People connector lift-and-shift guide (string + stringCollection)

This guide is **tool-agnostic**. It explains the exact requirements for a **people connector** when your schema only uses **string** and **stringCollection** properties (the same label set used in the demo). Use it to lift-and-shift existing people profile data into Microsoft Graph without changing your source system.

## 1) Lift-and-shift checklist

1. **Choose labels** from the supported list (section 4).
2. **Map source fields** into Graph person entities (section 5).
3. **Create the external connection** with `contentCategory = "people"` (section 3).
4. **Register the schema** using only string and stringCollection properties (section 3).
5. **Register the profile source** and update its precedence (section 7).
6. **Ingest items** with people-labeled property values encoded as JSON strings (section 6).

## 2) Graph API baseline for people connectors

- **Base URL:** `https://graph.microsoft.com/beta`
- People connectors require **Graph beta** for `contentCategory = "people"` and admin profile-source APIs.

## 3) Required connection and schema rules

These requirements are **non-negotiable** for a people connector:

**Connection**
- `contentCategory` must be **"people"**.
- `connectionId` must be **alphanumeric only** (no hyphens, underscores, or spaces).
- Provide a readable `name` and `description`.

**Schema**
- `baseType` must be **"microsoft.graph.externalItem"**.
- **Only** `string` and `stringCollection` property types are allowed for people labels.
- Property names must be **alphanumeric** and **<= 32 chars**.
- People connectors **do not support externalItem.content**.
- You must have **exactly one** property labeled **personAccount** (string).
- Do not use search flags for people-labeled properties (Graph ignores them).

## 4) Supported labels (string + stringCollection)

Use only the labels below. The **Type** column is the only allowed type for each label.

| Label | Type | Graph entity | Notes |
| --- | --- | --- | --- |
| personAccount | string | userAccountInformation | Required exactly once |
| personName | string | personName | |
| personCurrentPosition | string | workPosition | |
| personWebSite | string | personWebsite | |
| personNote | string | personAnnotation | |
| personAddresses | stringCollection | itemAddress | Max 3 entries |
| personEmails | stringCollection | itemEmail | Max 3 entries |
| personPhones | stringCollection | itemPhone | |
| personAwards | stringCollection | personAward | |
| personCertifications | stringCollection | personCertification | |
| personProjects | stringCollection | projectParticipation | |
| personSkills | stringCollection | skillProficiency | |
| personWebAccounts | stringCollection | webAccount | |
| personAnniversaries | stringCollection | personAnnualEvent | |

### Blocked labels (do not use)

- personManager
- personAssistants
- personColleagues
- personAlternateContacts
- personEmergencyContacts

## 5) Label payload rules (the key lift-and-shift requirement)

Each people-labeled property must be a **JSON-encoded Graph entity object**. That means:

- **string label**: value is a **JSON object string**.
- **stringCollection label**: value is an **array of JSON object strings**.

You must include **all required Graph fields** for each label's entity type. The required fields come from the Graph profile schema, so ensure your mapping includes those fields before ingesting.

Examples:

```json
// personName (string)
"{\"displayName\":\"Alex Wilber\"}"

// personSkills (stringCollection)
[
  "{\"displayName\":\"TypeScript\",\"proficiency\":\"advanced\"}",
  "{\"displayName\":\"Azure\",\"proficiency\":\"intermediate\"}"
]
```

## 6) External item payload requirements

People connectors require **everyone** access on each item and use **properties-only payloads**.

- `acl` must include `{ "type": "everyone", "value": "everyone", "accessType": "grant" }`.
- `properties` must contain **only** the people-labeled properties with JSON string values.
- Do not send `content` (or keep it empty if your client always includes it).

Example external item:

```json
{
  "id": "user@contoso.com",
  "acl": [
    { "type": "everyone", "value": "everyone", "accessType": "grant" }
  ],
  "properties": {
    "account": "{\"userPrincipalName\":\"user@contoso.com\"}",
    "displayName": "{\"displayName\":\"Jordan Kent\"}",
    "skills": [
      "{\"displayName\":\"TypeScript\",\"proficiency\":\"advanced\"}"
    ]
  }
}
```

## 7) Provisioning and profile source registration (Graph API)

Use these calls in order:

1. **Create external connection**
   - `POST /external/connections`
2. **Patch schema**
   - `PATCH /external/connections/{connectionId}/schema`
3. **Wait for schema ready**
   - `GET /external/connections/{connectionId}/schema`
4. **Register profile source**
   - `POST /admin/people/profileSources`
5. **List profile property settings**
   - `GET /admin/people/profilePropertySettings`
6. **Update precedence**
   - `PATCH /admin/people/profilePropertySettings/{id}`
7. **Ingest items**
   - `PUT /external/connections/{connectionId}/items/{itemId}`

Profile source payload example:

```json
{
  "sourceId": "peopleconnector",
  "displayName": "People Directory",
  "webUrl": "https://example.com/people"
}
```

Precedence payload example:

```json
{
  "prioritizedSourceUrls": [
    "https://graph.microsoft.com/beta/admin/people/profileSources(sourceId='peopleconnector')"
  ]
}
```

## 8) Demo-aligned minimal schema (string + stringCollection)

This is a minimal schema that mirrors the demo label set (string + stringCollection only):

```json
{
  "baseType": "microsoft.graph.externalItem",
  "properties": [
    { "name": "account", "type": "string", "labels": ["personAccount"] },
    { "name": "displayName", "type": "string", "labels": ["personName"] },
    { "name": "skills", "type": "stringCollection", "labels": ["personSkills"] }
  ]
}
```
