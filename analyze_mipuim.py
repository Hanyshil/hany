import pandas as pd
import re

df = pd.read_csv('/home/user/hany/mipuim_full.csv')

# שאלות ותיאוריהן
questions = {
    'Q1': 'איך תלמד חומר חדש שקשה להבין?',
    'Q2': 'מה תעשה כשאתה תקוע?',
    'Q3': 'איך תתמודד עם חומר ארוך/מורכב?',
    'Q4': 'מה המטרה שלך בלמידה?',
    'Q5': 'איך תדע אם למדת טוב?',
    'Q6': 'איך תרגיש אם תקבל ציון נמוך למרות שלמדת?',
    'Q7': 'למה קיבלת ציון נמוך?',
    'Q8': 'האם תלך לדבר עם המורה?',
    'Q9': 'מה תעשה לפני המבחן הבא?'
}

# סיווג תשובות לפי נושאים
def classify_ai_use(text):
    text = str(text).lower()
    if any(w in text for w in ['צאט','gpt','gpt','ai','בינה']):
        return 'שימוש ב-AI'
    return None

def classify_help_source(text):
    text = str(text)
    sources = []
    if any(w in text for w in ['צאט','gpt','GPT','AI']):
        sources.append('AI')
    if any(w in text for w in ['הורים','אמא','אבא','משפחה']):
        sources.append('הורים')
    if any(w in text for w in ['מורה','פרטית']):
        sources.append('מורה')
    if any(w in text for w in ['חבר','חברות','ילד']):
        sources.append('חברים')
    if any(w in text for w in ['אינטרנט','סרטון','יוטיוב']):
        sources.append('אינטרנט')
    if any(w in text for w in ['לא יודע','אין לי מושג']):
        sources.append('לא יודע')
    return sources if sources else ['עצמאי']

def classify_emotion(text):
    text = str(text)
    if any(w in text for w in ['עצוב','מבואס','אכזב','מאוכזב','עצבני','מתוסכל','בכה','תבאס']):
        return 'שלילי'
    if any(w in text for w in ['בסדר','כיף','טוב','אנסה','ישתפר']):
        return 'חיובי/ניטרלי'
    if any(w in text for w in ['לא יודע','לא ידעתי']):
        return 'לא יודע'
    return 'ניטרלי'

print("=" * 60)
print("ניתוח תשובות תלמידים - גיליון סימולציות עברית")
print("=" * 60)
print(f"\nסה\"כ תלמידים: {len(df)}")
print(f"בנים: {(df['מגדר']=='זכר').sum()} | בנות: {(df['מגדר']=='נקבה').sum()}")

# Q1 - שימוש ב-AI ללמידה
print("\n" + "="*60)
print("Q1 - איך תלמד חומר חדש שקשה להבין?")
print("-"*40)
ai_users = df[df['Q1'].str.contains('צאט|gpt|GPT|AI|בינה', case=False, na=False)]
internet_users = df[df['Q1'].str.contains('אינטרנט|סרטון', na=False)]
family_users = df[df['Q1'].str.contains('הורים|אמא|אבא|משפחה', na=False)]
dont_know = df[df['Q1'].str.contains('לא יודע|אין לי מושג', na=False)]

print(f"  שימוש ב-AI/צ'אט GPT:  {len(ai_users)} תלמידים ({len(ai_users)/len(df)*100:.0f}%)")
print(f"  אינטרנט/סרטונים:      {len(internet_users)} תלמידים ({len(internet_users)/len(df)*100:.0f}%)")
print(f"  עזרה ממשפחה:          {len(family_users)} תלמידים ({len(family_users)/len(df)*100:.0f}%)")
print(f"  לא יודע:              {len(dont_know)} תלמידים ({len(dont_know)/len(df)*100:.0f}%)")

# Q2 - מה תעשה כשתקוע
print("\n" + "="*60)
print("Q2 - מה תעשה כשאתה תקוע?")
print("-"*40)
ai_q2 = df[df['Q2'].str.contains('צאט|gpt|GPT|AI', case=False, na=False)]
teacher_q2 = df[df['Q2'].str.contains('מורה', na=False)]
family_q2 = df[df['Q2'].str.contains('הורים|אמא|אבא|משפחה', na=False)]
friend_q2 = df[df['Q2'].str.contains('חבר|חברות', na=False)]
print(f"  שימוש ב-AI/צ'אט GPT:  {len(ai_q2)} תלמידים ({len(ai_q2)/len(df)*100:.0f}%)")
print(f"  פנייה למורה:          {len(teacher_q2)} תלמידים ({len(teacher_q2)/len(df)*100:.0f}%)")
print(f"  פנייה למשפחה:         {len(family_q2)} תלמידים ({len(family_q2)/len(df)*100:.0f}%)")
print(f"  פנייה לחברים:         {len(friend_q2)} תלמידים ({len(friend_q2)/len(df)*100:.0f}%)")

# Q6 - תגובה רגשית לציון נמוך
print("\n" + "="*60)
print("Q6 - איך תרגיש אם תקבל ציון נמוך למרות שלמדת?")
print("-"*40)
neg_q6 = df[df['Q6'].str.contains('עצוב|מבואס|אכזב|מאוכזב|עצבני|מתוסכל|בכה|תבאס|אכזבה', na=False)]
pos_q6 = df[df['Q6'].str.contains('בסדר|אנסה|ישתפר|כיף', na=False)]
dk_q6 = df[df['Q6'].str.contains('לא יודע', na=False)]
print(f"  תגובה שלילית (כעס/עצב/אכזבה): {len(neg_q6)} תלמידים ({len(neg_q6)/len(df)*100:.0f}%)")
print(f"  תגובה חיובית/ניטרלית:          {len(pos_q6)} תלמידים ({len(pos_q6)/len(df)*100:.0f}%)")
print(f"  לא יודע:                        {len(dk_q6)} תלמידים ({len(dk_q6)/len(df)*100:.0f}%)")

# Q8 - פנייה למורה
print("\n" + "="*60)
print("Q8 - האם תלך לדבר עם המורה?")
print("-"*40)
yes_q8 = df[df['Q8'].str.contains('אלך|יבוא|כן|ילך', na=False)]
no_q8 = df[df['Q8'].str.contains('לא|לא יודע', na=False)]
print(f"  כן - יפנה למורה:   {len(yes_q8)} תלמידים ({len(yes_q8)/len(df)*100:.0f}%)")
print(f"  לא/לא יודע:        {len(no_q8)} תלמידים ({len(no_q8)/len(df)*100:.0f}%)")

# Q9 - תוכנית ללמידה
print("\n" + "="*60)
print("Q9 - מה תעשה לפני המבחן הבא?")
print("-"*40)
study_more = df[df['Q9'].str.contains('ילמד|אלמד|חזרה|תרגול', na=False)]
private = df[df['Q9'].str.contains('מורה פרטית|מורה פרטי', na=False)]
dk_q9 = df[df['Q9'].str.contains('לא יודע|אין לי מושג', na=False)]
print(f"  ילמד יותר/יחזור על החומר: {len(study_more)} תלמידים ({len(study_more)/len(df)*100:.0f}%)")
print(f"  מורה פרטית:               {len(private)} תלמידים ({len(private)/len(df)*100:.0f}%)")

print("\n" + "="*60)
print("סיכום ממצאים מרכזיים")
print("="*60)

