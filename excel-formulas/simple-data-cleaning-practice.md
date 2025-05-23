# **🔥 Excel URL Extraction Tutorial: From Messy Data to Clean Links 🔥**

**Problem:** You pasted a huge block of text (URLs mixed with junk) into Excel and need to extract ONLY HTTPS links into a clean vertical list.

**Solution:** A powerful combo of Excel's TEXT-TO-COLUMNS + FILTER + TRANSPOSE functions!

---

## **🛠️ STEP-BY-STEP GUIDE**

### **📌 ORIGINAL DATA**

(Copy and paste the below data to your Excel cell A1)

| A                           |B|C|
|-----------------------------|--|--|
| {junk data} https://site1.com more junk https://site2.com url https://site3.com

---

## **🔧 STEP 1: SPLIT BY SPACES (Text-to-Columns)**

### **1️⃣ Separate Data Using Text-to-Columns** (Convert wizard🧙‍♂️)
**(When everything is pasted into one cell, but needs to be split into others.)**

**Manual Method:**
1. Select cell A1 (with your pasted data)
2. Go to: **Data → Text to Columns → Delimited → Next** 
3. Check **"Space"** as delimiter → Finish
4. Your data is now split across columns (A1:H1) 

😎

**What happens:**  
Excel separates your long text into individual cells wherever spaces occur.   


**Result:**
### **Version 2: Parentheses** 
| A       | B      | C                  | D     | E     | F                  | G    | H                  |
|---------|--------|--------------------|-------|-------|--------------------|------|--------------------|
| {junk   | data}  | https://site1.com  | more  | junk  | https://site2.com  | url  | https://site3.com  |


---
## **2️⃣ STEP 2: Extract ONLY HTTPS URLs**
*(Using FILTER magic)* ☕✨

 Now lets click in __cell A2__ and paste the formula above:  

**Formula:**
```excel
=FILTER(A1:H1; LEFT(A1:H1; 5)="https")
```

### **🧠 FORMULA BREAKDOWN:**
- `A1:F1` → The range to check (all split cells)
- `LEFT(text, num_chars)` → Extracts characters from the start  
  - `LEFT(A1:H1, 5)` → Takes first 5 characters of each cell
  - `="https"` → Checks if those 5 chars are "https"  

>[!NOTE]  
>**What it does:** Scans all cells and returns ONLY those starting with "https"  
>**Output:** Horizontal list of clean URLs  

**Result:**  
|A|B|C|
|---------|--------|--------------------|
https://site1.com | https://site2.com | https://site3.com|

*(Horizontal output in multiple cells)* ✨ ↔️ ✨

---

## **3️⃣ STEP 3: Flip Results with TRANSPOSE to vertical column ↕️** 
*(TRANSPOSE is gold!)* 🏅🪙

Lets part, click in cell __A3__  and paste the above formula:

**Formula:**  

```excel
=TRANSPOSE(FILTER(A2:C2; LEFT(A2:C2; 5)="https"))
```

**What TRANSPOSE does:**  

Converts horizontal ranges to vertical (or vice versa) ↕️

**Final Output:**  

| A             |
|---------------|
| https://site1.com |
| https://site2.com |
| https://site3.com |

---

## **🎯 VISUALIZATION OF THE PROCESS**
```
Original Text
│
├─ Text-to-Columns (Split by space)🧙‍♂️
│  → {junk data} | https://site1.com | more junk | etc...
│
├─ FILTER(LEFT;5 = "https") ☕
│  → https://site1.com | https://site2.com | https://site3.com
│
└─ TRANSPOSE 🪙
   ↓
   https://site1.com
   https://site2.com
   https://site3.com
```

---

## **💡 PRO TIPS**
1. **Handle Empty Cells:**  
   ```excel
   =TRANSPOSE(FILTER(A1:F1, (LEFT(A1:F1,5)="https")*(A1:F1<>""))
   ```

2. **Make Links Clickable:**  
   ```excel
   =HYPERLINK(TRANSPOSE(FILTER(A1:F1, LEFT(A1:F1,5)="https")))
   ```

3. **Count Extracted URLs:**  
   ```excel
   =COUNTA(TRANSPOSE(FILTER(A1:F1, LEFT(A1:F1,5)="https")))
   ```

---

## **🚨 COMMON ISSUES & FIXES**
- 1️⃣ **Problem:** `Error`   
  **Solution:** The correct punctuation for formulas depends on the version of Excel you are using, try `;` or `,`    
     
 __For Example:__  
  `Change`   
  ```excel
  =TRANSPOSE(FILTER(A2:C2; LEFT(A2:C2; 5)="https"))
  ```  
  `For`  
  ```excel
  =TRANSPOSE(FILTER(A2:C2, LEFT(A2:C2, 5)="https"))
  ```  

- 2️⃣ **Problem:** Mixed http/https links  
  **Solution:** Catch both:  
  ```excel
  =TRANSPOSE(FILTER(A1:F1, (LEFT(A1:F1,4)="http")))
  ```

---

## **📊 FINAL RESULT SHOWCASE**
| Extracted URLs       |
|----------------------|
| https://site1.com    |
| https://site2.com    |
| https://site3.com    |

I hope this tutorial helped you understand how to clean data, bye, bye ✨ 3= "🧙‍♂️"  

#### Official Microsoft reference:
__Transpose:__ https://support.microsoft.com/en-us/office/transpose-function-ed039415-ed8a-4a81-93e9-4b6dfac76027  
__Filter:__ https://support.microsoft.com/en-us/office/filter-function-f4f7cb66-82eb-4767-8f7c-4877ad80c759

Tutorial by:  
🤵🏽[LinkedIn/thiago-cequeira-99202239](https://www.linkedin.com/in/thiago-cequeira-99202239/) \
🤗[huggingface.co/ThiSecur](https://huggingface.co/ThiSecur)

