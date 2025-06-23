# **PowerPoint Presentation Generator - VBA Automation Tool**  

## **📌 Overview**  
This **VBA macro** automates the creation of PowerPoint presentations with:  
✅ **Title Slide** (First slide)  
✅ **Index Slide** (Second slide with topics)  
✅ **Custom Slides** (User-defined number of slides)  

![image](https://github.com/user-attachments/assets/ea01ed0a-1e6e-45c7-87ce-753320fad450)  
![image](https://github.com/user-attachments/assets/6a520152-0fa2-47fe-a783-6b9d8c5ddfaf)
![image](https://github.com/user-attachments/assets/afd16c5d-1d48-48a0-bcc3-4073bbc5922a)  

Slides Order:  
     - **Slide 1**: Title Slide (Main Title)   
     - **Slide 2**: Index (List of Topics)   
     - **Slides 3, 4+**: Custom slides (user inputs content via InputBox)    

>[!Tip]
> - You have to add the **UserForm** for easy input.     
> - See the **Deployment Guide** below.  

---

## **🛠️ Deployment Guide**  

### **Step 1: Open VBA Editor in PowerPoint**  
1. Open **PowerPoint** → **Blank Presentation**.  
2. Press **`Alt + F11`** to open the **VBA Editor**.  
3. (If Developer tab is missing: **File → Options → Customize Ribbon → Check "Developer" → OK**).  

### **Step 2: Import the Code**  
1. **Insert a UserForm**:  
   - Right-click in **Project Explorer** → **Insert → UserForm**.  
   - Rename it to **`UserForm1`** (if not default).  

2. **Add Controls to the UserForm**:  
   - Open the **Toolbox** (`View → Toolbox`).  
   - Add these controls with **exact names**:  

   | Control Type | Name | Purpose |
   |-------------|------|---------|
   | **TextBox** | `txtMainTitle` | Presentation title |
   | **TextBox** | `txtNumSlides` | Number of slides (minimum 2) |
   | **TextBox** | `txtTopicName` | Topic name input |
   | **ListBox** | `lstTopics` | Displays added topics |
   | **CommandButton** | `cmdAddTopic` | Adds topic to list |
   | **CommandButton** | `cmdRemoveTopic` | Removes selected topic |
   | **CommandButton** | `cmdGenerate` | Creates the presentation |
   | **CommandButton** | `cmdCancel` | Closes the form |

3. **Paste the Code**:  
   - Double-click the **UserForm** → Paste the **UserForm code**.  
   - Insert a **Module** (`Insert → Module`) → Paste the **Standard Module code**.  

### **Step 3: Run the Macro**  
1. **Save as Macro-Enabled Presentation** (`File → Save As → .pptm`).  
2. **Enable Macros** when prompted (if security warning appears).  
3. Run the macro:  
   - Press **`Alt + F8`** → Select **`ShowPresentationGenerator`** → **Run**.  
   - Or, assign it to a **button** (`Developer → Insert → Button`).  

---

## **🎯 How It Works**  
1. **User Input**  
   - Enter **presentation title**, **number of slides**, and **topics**.  
   - Click **"Add Topic"** to populate the list.  

2. **Generate Presentation**  
   - Click **"Generate"** → PowerPoint creates:  
     - **Slide 1**: Title Slide  
     - **Slide 2**: Index (List of Topics)  
     - **Slides 3+**: Custom slides (user inputs content via InputBox)  

3. **Save the Presentation**  
   - A **Save As dialog** opens → Choose **format (PPTX, PPTM, PPT)**.  

---

## **⚠️ Troubleshooting**  
| Issue | Solution |
|-------|----------|
| **Macros not running** | Enable macros in **Trust Center Settings** (`File → Options → Trust Center → Macro Settings → Enable All Macros`). |
| **"Type Mismatch" error** | Ensure `txtNumSlides` has a **valid number (≥ 2)**. |
| **FileDialog not working** | Check if `Microsoft Office XX.X Object Library` is enabled (`Tools → References`). |

---

## **📥 Download VBA Scripts/Code**  
🔗 **[UserForm Code](https://github.com/ThiagoMaria-SecurityIT/Microsoft_Office_Automation_VBA/blob/main/VBA/PowerPoint/UserForm-Code.vba)**  
🔗 **[Module Code](https://github.com/ThiagoMaria-SecurityIT/Microsoft_Office_Automation_VBA/blob/main/VBA/PowerPoint/Module-Code.vba)**  
 

---

## **📜 License**  
MIT License - Free for personal/commercial use.  

---

### **🚀 Ready to Automate PowerPoint?**  
Follow the steps above, and you’ll generate professional presentations in seconds!   

---  

## About the Author   

**Thiago Maria - From Brazil to the World 🌎**  
*Senior Security Information Professional | Passionate Programmer | AI Developer*

With a professional background in security analysis and a deep passion for programming, I created this repo share some knowledge about security information, cybersecurity, Python and development practices. Most of my work here focuses on implementing security-first approaches in developer tools while maintaining usability.

Lets Connect:

[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-blue)](https://www.linkedin.com/in/thiago-cequeira-99202239/)  
[![Hugging Face](https://img.shields.io/badge/🤗Hugging_Face-AI_projects-yellow)](https://huggingface.co/ThiSecur)

## Ways to Contribute:   
✨ **Contributions welcome!** Fork → Improve → Submit a **Pull Request**.   
✨ Want to see more upgrades? Help me keep it updated!    
 [![Sponsor](https://img.shields.io/badge/Sponsor-%E2%9D%A4-red)](https://github.com/sponsors/ThiagoMaria-SecurityIT) 
