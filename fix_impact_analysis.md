# AutoNQ AI — Image Ownership Fix Impact Analysis

This document provides a highly rigorous, comprehensive **Fix Impact Analysis, Risk Assessment, and Dependency Analysis** regarding the proposed changes to the image ownership system. 

---

## 1. THE COMPLETE LIFECYCLE OF `image_base64`

The following lifecycle traces how image data flows through the application's components:

```
[Phase 1: Upload / Capture] (app.py)
   st.file_uploader / st.camera_input 
   └─> Captured as raw bytes (bytes/bytearray) in memory
   └─> Stored in st.session_state.observations list as {"text": ..., "image": bytes}

[Phase 2: Save Submission] (app.py -> excel_backend.py)
   Frontend submits observations loop:
   └─> Encodes raw bytes to base64 string (image_b64)
   └─> Calls add_audit_entry({"image_base64": image_b64, ...})
   Backend add_audit_entry normalizes data:
   └─> [BUG 1] Checks entry_data.get("image") instead of "image_base64" (mismatch)
   └─> Resolves to "" and writes blank cell to workbook

[Phase 3: Database Storage] (data/audit_master_data.xlsx)
   Stored physically in the 19th column ("image_base64") of the "audit_entries" sheet.

[Phase 4: Retrieval & Normalization] (excel_backend.py -> app.py)
   Backend loads sheet via load_audit_entries():
   └─> Centralized normalize_columns() runs
   └─> inject_aliases() runs
   └─> app.py:378-390 maps "image" column to "image_base64" in full_data if present

[Phase 5: Q-Check Checklist Generation] (app.py -> setup_environment.py)
   st.button("Generate Q-Check"):
   └─> Passes records from full_data to generate_qcheck_questions()
   └─> [BUG 3] AI generator strips "image_base64" column and sets "Ref Photo" to ""
   └─> Returning checklist records do not carry images

[Phase 6: Frontend Joining Fallback] (app.py)
   Streamlit page rendering:
   └─> Frontend builds _img_lookup dictionary keyed by Station Name (Station -> Latest Image)
   └─> [BUG 3] Injects the station-level image into all checkpoints sharing that station
   └─> Decodes base64 back to image bytes and displays via st.image()

[Phase 7: Q-Check NOK Save] (app.py -> excel_backend.py)
   st.button("Submit Q-Check"):
   └─> Loops through NOK checklist records and calls add_audit_entry()
   └─> [BUG 2] Passes "image_base64": image_b64 (dangling global module-level variable)
   └─> Overwrites database rows with unrelated images from the Audit Entry page
```

---

## 2. IMAGE PATHWAYS MATRIX

| File Path | Code Block / Lines | Operation Type | Variable / Column Name | Context / Role |
| :--- | :--- | :--- | :--- | :--- |
| [app.py](file:///d:/HACK28/app.py) | `L610-619` | **Created** | `image_bytes` | Captures raw image bytes from uploader/camera and stores them in memory state. |
| [app.py](file:///d:/HACK28/app.py) | `L663-671` | **Created / Encoded** | `image_b64` | Encodes raw bytes to base64 string for API payload. |
| [app.py](file:///d:/HACK28/app.py) | `L687` | **Passed to Save** | `"image_base64"` | Key in the dict payload sent to the backend. |
| [excel_backend.py](file:///d:/HACK28/excel_backend.py) | `L47` | **Schema Schema** | `"image_base64"` | Column name in `AUDIT_COLUMNS` database schema. |
| [excel_backend.py](file:///d:/HACK28/excel_backend.py) | `L430-438` | **Modified / Read** | `entry_data.get("image")` | Backend attempts to parse the payload (key mismatch bug). |
| [excel_backend.py](file:///d:/HACK28/excel_backend.py) | `L460, L487-488` | **Stored** | `image_b64` | Writes final base64 string to cells in sheet. |
| [excel_backend.py](file:///d:/HACK28/excel_backend.py) | `L222-250` | **Retrieved** | `load_audit_entries()` | Reads physical sheet rows from disk into memory DataFrame. |
| [app.py](file:///d:/HACK28/app.py) | `L378-390` | **Modified** | `full_data["image_base64"]` | Frontend column alignment utility. |
| [app.py](file:///d:/HACK28/app.py) | `L1084-1105` | **Modified (lossy)** | `_img_lookup` | Fallback station-to-image mapping that causes repeated images. |
| [app.py](file:///d:/HACK28/app.py) | `L1131-1150` | **Rendered** | `row.get("image_base64")` | Decodes base64 string and renders image preview inside Streamlit column. |
| [app.py](file:///d:/HACK28/app.py) | `L1183` | **Stored (leak)** | `"image_base64": image_b64` | Leakage of global variable into new database rows. |

---

## 3. DESIGN DECISION: IMAGE OWNERSHIP PROVENANCE

We must mathematically define where image ownership belongs in a manufacturing execution and auditing context:

```
Assembly Line  (Line 1)
   └── Station  (Station A)
        └── Checklist Question / Checkpoint  (Yes/No Checkpoint)
             └── Deviation Event / Observation  (A specific physical failure at a point in time)
                  └── Evidence Photograph  (The exact snapshot of the failure)
```

### Provenance Evaluation:
1. **Assembly Line / Station**: Ownership **cannot** reside here. A station is a static physical structure. An image represents a dynamic deviation event that occurs at a specific point in time. Storing images at the station level results in historical errors leaking into new checks.
2. **Question / Checkpoint**: Ownership **cannot** reside here. An AI-generated question is a generic rule ("Is the safety guard closed?"). The same question can be verified daily; however, each daily check will have its own unique physical state.
3. **Audit Entry Row / Observation**: Ownership **MUST** reside here. The image represents the physical evidence of a **specific, single deviation occurrence**. It is tied 1-to-1 to that audit record.
4. **Q-Check Reference Photo**: When a historical deviation is converted into a Q-Check checkpoint, the Q-Check item should display the **exact evidence photo from that historical row** to show the auditor what the previous deviation looked like.

**Verdict**: Image ownership belongs exclusively to the **Audit Entry Row (Observation)**. The Q-Check page must carry this relationship downstream 1-to-1 rather than flattening it to the station level.

---

## 4. CURRENT DATABASE SCHEMA

The exact database schema currently implemented in `excel_backend.py:42-48` is as follows:

| Col # | Excel Header Name | Data Type | Purpose |
| :--- | :--- | :--- | :--- |
| 1 | `audit_id` | String / UUID | Unique primary key for the audit entry |
| 2 | `audit_date` | Date / DateTime | The date the audit was performed |
| 3 | `audit_time` | String (HH:MM) | The time the audit was performed |
| 4 | `line` | String | Production line name (e.g. HVML) |
| 5 | `area` | String | Factory department or area |
| 6 | `station_no` | String | Dynamic station identifier code |
| 7 | `station_name` | String | Human-readable station name |
| 8 | `checkpoint` | String | The audit principle / checkpoint being checked |
| 9 | `expected_result` | String | The standard standard operating procedure |
| 10 | `actual_result` | String | The actual finding / observation description |
| 11 | `remarks` | String | Optional remarks or comments |
| 12 | `severity` | String | Severity rating (Low, Medium, High, Critical) |
| 13 | `category` | String | Principle Category |
| 14 | `auditor_name` | String | Name of the auditing personnel |
| 15 | `flm_name` | String | Front Line Manager / Supervisor name |
| 16 | `shift` | String | Working shift (Shift 1, Shift 2, Shift 3) |
| 17 | `status` | String | Deviation status (Open, Closed) |
| 18 | `created_at` | DateTime | Timestamp when database row was written |
| 19 | `image_base64` | String (Long text) | **Base64 string representing the unique observation image** |

---

## 5. FORENSIC VALIDATION QUESTIONS

### Q-1: Can multiple questions legitimately share the same image?
* **No**. In a real manufacturing plant, different questions represent distinct checkpoints (e.g., "Is the operator wear ESD shoes?" vs "Is the torque wrench calibrated?"). Linking the same torque wrench photo to the ESD shoe check is incorrect, confusing, and violates audit standards. Images must remain 1-to-1 with the specific observation row.

### Q-2: Is preserving `image_base64` inside `generate_qcheck_questions()` sufficient?
* **Yes, absolutely**. If the AI generator maps the original `image_base64` to the returned question dictionary directly, each checkpoint will carry its own unique reference image. The frontend can then render it directly, bypassing the lossy station-level lookup block.

### Q-3: Do any other dashboards, reports, or analytics depend on the station-level image logic?
* **None**. Ripgrep/Grep verification across the entire project confirms that `_img_lookup` is **exclusively** used inside the Daily Q-Check rendering loop in `app.py`. No dashboards, reports, or background tasks rely on it. It can be safely removed.

---

## 6. SPECIFIC MODIFICATIONS REQUIRED

The following files and precise lines require surgical modification to resolve the issues:

### 1. `excel_backend.py` (Line 430)
* **Current**:
  ```python
  image_raw = entry_data.get("image", "")
  ```
* **Change to**:
  ```python
  image_raw = entry_data.get("image_base64", entry_data.get("image", ""))
  ```
* **Impact**: Ensures that both `"image_base64"` (sent by app.py) and `"image"` keys are correctly read, fixing the saving of empty cells.

### 2. `setup_environment.py` (Lines 755-758)
* **Current**:
  ```python
  questions.append({
      "Station": row["station"], "Checkpoint": q,
      "Ref Photo": "", "Status": "OK", "Remark": ""
  })
  ```
* **Change to**:
  ```python
  questions.append({
      "Station": row["station"], "Checkpoint": q,
      "Ref Photo": "", "Status": "OK", "Remark": "",
      "image_base64": row.get("image_base64", "")
  })
  ```
* **Impact**: Instructs the AI-checklist generator to preserve the specific, unique image from the source deviation row.

### 3. `app.py` (Lines 1079-1106)
* **Current**:
  ```python
  # ── FIX-B: Re-join image_base64 from full_data into Q-Check records ──
  _img_lookup: dict = {}
  ...
  for _rec in records:
      if not _rec.get("image_base64"):
          ...
          _rec["image_base64"] = _img_lookup.get(_rec_station, "")
  ```
* **Change to**:
  * **Remove this block entirely**.
* **Impact**: Eliminates the lossy station-based image fallback logic since images are now correctly populated directly within the `records` payload.

### 4. `app.py` (Line 1183)
* **Current**:
  ```python
  "image_base64": image_b64,
  ```
* **Change to**:
  ```python
  "image_base64": "",
  ```
* **Impact**: Solves the Dangling Global Reference leak by ensuring new, fresh NOK findings are logged cleanly (without carrying over old images from other pages).

---

## 7. RISK ASSESSMENT & MITIGATION Strategy

| Identified Risk | Severity | Potential Impact | Mitigation Strategy |
| :--- | :--- | :--- | :--- |
| **Missing Image Data in Legacy Records** | Low | Checkpoints might render without photos. | Fallback gracefully in the UI. If `image_base64` is empty, display the standard `"No Image"` caption. |
| **NameError on Q-Check Submission** | High | App crashes when clicking "Submit Q-Check" due to missing variable. | Surgical removal of the global `image_b64` reference from the Q-Check save dictionary, replacing it with `""`. |
| **Excel Cell Write Bloat** | Medium | Saving large Base64 strings could slow down Excel writes. | The thread-safe writing mechanism (`_WRITE_LOCK`) and caches are already optimized for this. The base64 sanitization on read ensures smooth decoding. |

---

## 8. SAFE PATCH PLAN

We will execute the fixes in the following structured sequence:

```
[Step 1: Backend Integration] ──> Update excel_backend.py:430 to handle both keys.
                                    (Restores image save to Excel database)
                                     │
[Step 2: Core Data Preservation] ──> Update setup_environment.py:755 to include image_base64.
                                    (Maintains 1-to-1 question-to-image link)
                                     │
[Step 3: UI Layer Alignment] ────> Remove _img_lookup fallback block from app.py:1079.
                                    (Eliminates repeated station-level images)
                                     │
[Step 4: Leak Prevention] ───────> Set "image_base64" to "" in app.py:1183.
                                    (Prevents dangling global reference leak & NameError)
```

---

## 9. FINAL RECOMMENDATION

The proposed plan is **highly safe, targeted, and complete**. It directly addresses the root cause of the bug by restoring data integrity at the storage and transformation layers instead of patching it loosely on the UI layer. 

No code will be written yet. We are awaiting your explicit approval to execute this Safe Patch Plan.
