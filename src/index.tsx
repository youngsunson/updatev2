import { useState, useEffect } from 'react';
import ReactDOM from 'react-dom/client';
import './index.css';

/* -------------------------------------------------------------------------- */
/*                                TYPES                                       */
/* -------------------------------------------------------------------------- */

interface Correction {
  wrong: string;
  suggestions: string[];
  position?: number;
}

interface ToneSuggestion {
  current: string;
  suggestion: string;
  reason: string;
}

interface StyleSuggestion {
  current: string;
  suggestion: string;
  type: string;
}

interface StyleMixing {
  detected: boolean;
  recommendedStyle?: string;
  reason?: string;
  corrections?: Array<{
    current: string;
    suggestion: string;
    type: string;
  }>;
}

interface PunctuationIssue {
  issue: string;
  currentSentence: string;
  correctedSentence: string;
  explanation: string;
}

interface EuphonyImprovement {
  current: string;
  suggestions: string[];
  reason: string;
}

interface ContentAnalysis {
  contentType: string;
  description?: string;
  missingElements?: string[];
  suggestions?: string[];
}

/* -------------------------------------------------------------------------- */
/*                        PROMPT BUILDERS                                     */
/* -------------------------------------------------------------------------- */

const buildTonePrompt = (text: string, tone: string) => {
  const toneInstructions: Record<string, string> = {
    'formal': `আপনি একজন বাংলা ভাষা বিশেষজ্ঞ। নিচের টেক্সটকে **আনুষ্ঠানিক (Formal)** টোনে রূপান্তরের জন্য বিশ্লেষণ করুন। বৈশিষ্ট্য: আপনি/আপনার ব্যবহার, ক্রিয়াপদে 'করুন/বলুন', পূর্ণ বাক্য গঠন।`,
    'informal': `আপনি একজন বাংলা ভাষা বিশেষজ্ঞ। নিচের টেক্সটকে **অনানুষ্ঠানিক (Informal)** টোনে রূপান্তরের জন্য বিশ্লেষণ করুন। বৈশিষ্ট্য: তুমি/তুই ব্যবহার, কথ্য ভাষা, সহজ শব্দ।`,
    'professional': `আপনি একজন বাংলা ভাষা বিশেষজ্ঞ। নিচের টেক্সটকে **পেশাদার (Professional)** টোনে রূপান্তরের জন্য বিশ্লেষণ করুন। বৈশিষ্ট্য: স্পষ্টতা, আত্মবিশ্বাসী ভাষা, পেশাদার শব্দভাণ্ডার।`,
    'friendly': `আপনি একজন বাংলা ভাষা বিশেষজ্ঞ। নিচের টেক্সটকে **বন্ধুত্বপূর্ণ (Friendly)** টোনে রূপান্তরের জন্য বিশ্লেষণ করুন। বৈশিষ্ট্য: উষ্ণ সম্বোধন, আবেগপূর্ণ শব্দ, ইতিবাচক ভাষা।`,
    'respectful': `আপনি একজন বাংলা ভাষা বিশেষজ্ঞ। নিচের টেক্সটকে **সম্মানজনক (Respectful)** টোনে রূপান্তরের জন্য বিশ্লেষণ করুন। বৈশিষ্ট্য: সম্মানসূচক সম্বোধন, বিনীত অনুরোধ, শ্রদ্ধাসূচক শব্দ।`,
    'persuasive': `আপনি একজন বাংলা ভাষা বিশেষজ্ঞ। নিচের টেক্সটকে **প্রভাবশালী (Persuasive)** টোনে রূপান্তরের জন্য বিশ্লেষণ করুন। বৈশিষ্ট্য: শক্তিশালী শব্দ, জরুরিতা তৈরি, ইতিবাচক ফলাফল।`,
    'neutral': `আপনি একজন বাংলা ভাষা বিশেষজ্ঞ। নিচের টেক্সটকে **নিরপেক্ষ (Neutral)** টোনে রূপান্তরের জন্য বিশ্লেষণ করুন। বৈশিষ্ট্য: বস্তুনিষ্ঠ ভাষা, আবেগমুক্ত শব্দ, সূত্র উল্লেখ।`,
    'academic': `আপনি একজন বাংলা ভাষা বিশেষজ্ঞ। নিচের টেক্সটকে **শিক্ষামূলক (Academic)** টোনে রূপান্তরের জন্য বিশ্লেষণ করুন। বৈশিষ্ট্য: পরিভাষা ব্যবহার, তৃতীয় পুরুষ, জটিল বাক্য।`
  };

  return `${toneInstructions[tone]}

📝 **বিশ্লেষণের জন্য টেক্সট:**
"${text}"

📋 **আপনার কাজ:**
1. টেক্সটের প্রতিটি শব্দ ও বাক্যাংশ বিশ্লেষণ করুন।
2. কাঙ্ক্ষিত টোনে নেই এমন শব্দগুলো চিহ্নিত করুন।
3. **গুরুত্বপূর্ণ:** "current" ফিল্ডে শব্দটি হুবহু ইনপুট টেক্সট থেকে কপি করবেন (কোনো পরিবর্তন ছাড়া)।

📤 **Response Format (JSON only):**
{
  "toneConversions": [
    {
      "current": "বর্তমান শব্দ (হুবহু টেক্সট থেকে)",
      "suggestion": "সংশোধিত রূপ",
      "reason": "কারণ"
    }
  ]
}`;
};

const buildStylePrompt = (text: string, style: string) => {
  const styleInstructions: Record<string, string> = {
    'sadhu': `নিচের টেক্সটকে **সাধু রীতি**তে রূপান্তরের জন্য বিশ্লেষণ করুন। ক্রিয়াপদ (ছি->তেছি, ল->ইল), সর্বনাম (তার->তাহার) এবং অব্যয় পরিবর্তন করুন।`,
    'cholito': `নিচের টেক্সটকে **চলিত রীতি**তে রূপান্তরের জন্য বিশ্লেষণ করুন। ক্রিয়াপদ (তেছি->ছি, ইল->ল), সর্বনাম (তাহার->তার) এবং অব্যয় পরিবর্তন করুন।`
  };

  return `${styleInstructions[style]}

═══════════════════════════════════════
📝 বিশ্লেষণের জন্য টেক্সট:
"${text}"
═══════════════════════════════════════

⚠️ **সতর্কতা:**
- "current" ফিল্ডে শব্দটি টেক্সট থেকে **হুবহু কপি** করবেন।
- যদি কোন শব্দ পরিবর্তন প্রয়োজন না হয় তবে সেটি বাদ দিন।

📤 **Response Format (JSON only):**
{
  "styleConversions": [
    {
      "current": "বর্তমান শব্দ (হুবহু টেক্সট থেকে)",
      "suggestion": "সংশোধিত শব্দ",
      "type": "ক্রিয়াপদ/সর্বনাম/অব্যয়"
    }
  ]
}`;
};

/* -------------------------------------------------------------------------- */
/*                           MAIN COMPONENT                                   */
/* -------------------------------------------------------------------------- */

function App() {
  // Settings State
  const [apiKey, setApiKey] = useState(localStorage.getItem('gemini_api_key') || '');
  const [selectedModel, setSelectedModel] = useState(localStorage.getItem('gemini_model') || 'gemini-2.0-flash');
  
  // UI State
  const [isLoading, setIsLoading] = useState(false);
  const [loadingText, setLoadingText] = useState('');
  const [message, setMessage] = useState<{ text: string; type: 'success' | 'error' } | null>(null);
  const [activeModal, setActiveModal] = useState<'none' | 'settings' | 'instructions' | 'tone' | 'style'>('none');
  
  // Selection State
  const [selectedTone, setSelectedTone] = useState('');
  const [selectedStyle, setSelectedStyle] = useState('none');

  // Data State
  const [corrections, setCorrections] = useState<Correction[]>([]);
  const [toneSuggestions, setToneSuggestions] = useState<ToneSuggestion[]>([]);
  const [styleSuggestions, setStyleSuggestions] = useState<StyleSuggestion[]>([]);
  const [languageStyleMixing, setLanguageStyleMixing] = useState<StyleMixing | null>(null);
  const [punctuationIssues, setPunctuationIssues] = useState<PunctuationIssue[]>([]);
  const [euphonyImprovements, setEuphonyImprovements] = useState<EuphonyImprovement[]>([]);
  const [contentAnalysis, setContentAnalysis] = useState<ContentAnalysis | null>(null);
  
  const [stats, setStats] = useState({ totalWords: 0, errorCount: 0, accuracy: 100 });

  useEffect(() => {
    // Initialize Office
  }, []);

  /* --- HELPERS --- */
  const showMessage = (text: string, type: 'success' | 'error') => {
    setMessage({ text, type });
    setTimeout(() => setMessage(null), 3000);
  };

  const saveSettings = () => {
    localStorage.setItem('gemini_api_key', apiKey);
    localStorage.setItem('gemini_model', selectedModel);
    showMessage('সেটিংস সংরক্ষিত হয়েছে! ✓', 'success');
    setActiveModal('none');
  };

  /* --- WORD API INTERACTION --- */
  
  // IMPROVEMENT 1: Selection vs Whole Document Check
  const getTextFromWord = async (): Promise<string> => {
    return new Promise((resolve) => {
      Word.run(async (context) => {
        // Check for selection first
        const selection = context.document.getSelection();
        selection.load(['text', 'isEmpty']);
        await context.sync();

        let targetText = '';

        if (!selection.isEmpty && selection.text.trim().length > 0) {
          // User has selected text
          targetText = selection.text;
        } else {
          // No selection, get whole body
          const body = context.document.body;
          body.load('text');
          await context.sync();
          targetText = body.text;
        }
        
        // Normalize newlines to help AI understand structure
        const cleanText = targetText.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
        resolve(cleanText);
      }).catch((error) => {
        console.error('Error reading Word:', error);
        resolve('');
      });
    });
  };

  const highlightInWord = async (text: string, color: string) => {
    const cleanText = text.trim();
    if (!cleanText) return;

    await Word.run(async (context) => {
      // Search logic
      const results = context.document.body.search(cleanText, { 
        matchCase: false, 
        matchWholeWord: false, 
        ignoreSpace: true 
      });
      results.load('font');
      await context.sync();
      
      for (let i = 0; i < results.items.length; i++) {
        results.items[i].font.highlightColor = color;
      }
      await context.sync();
    }).catch(console.error);
  };

  // IMPROVEMENT 3: Safer Replacement with matchWholeWord & State Update
  const replaceInWord = async (oldText: string, newText: string) => {
    const cleanOldText = oldText.trim();
    let success = false;

    await Word.run(async (context) => {
      // Use matchWholeWord: true to avoid partial replacements (e.g. 'ban' inside 'band')
      const results = context.document.body.search(cleanOldText, { 
        matchCase: true, 
        matchWholeWord: true, 
        ignoreSpace: true 
      });
      results.load('items');
      await context.sync();

      if (results.items.length > 0) {
        results.items.forEach((item) => {
          item.insertText(newText, Word.InsertLocation.replace);
          item.font.highlightColor = "None";
        });
        await context.sync();
        success = true;
      }
    }).catch(console.error);

    if (success) {
      // Update UI State accurately
      const isNotMatch = (textToCheck: string) => textToCheck.trim() !== cleanOldText;

      setCorrections(prev => prev.filter(c => isNotMatch(c.wrong)));
      setToneSuggestions(prev => prev.filter(t => isNotMatch(t.current)));
      setStyleSuggestions(prev => prev.filter(s => isNotMatch(s.current)));
      setEuphonyImprovements(prev => prev.filter(e => isNotMatch(e.current)));
      setPunctuationIssues(prev => prev.filter(p => isNotMatch(p.currentSentence))); 
      
      setLanguageStyleMixing(prev => {
        if (!prev || !prev.corrections) return prev;
        const filtered = prev.corrections.filter(c => isNotMatch(c.current));
        return filtered.length > 0 ? { ...prev, corrections: filtered } : null;
      });

      showMessage(`সংশোধিত হয়েছে ✓`, 'success');
    } else {
      showMessage(`শব্দটি খুঁজে পাওয়া যায়নি (অন্য কোথাও পরিবর্তিত হতে পারে)।`, 'error');
    }
  };

  // IMPROVEMENT 2: Dismiss/Ignore Function
  const dismissSuggestion = (type: 'spelling' | 'tone' | 'style' | 'mixing' | 'punct' | 'euphony', textToDismiss: string) => {
    const cleanText = textToDismiss.trim();
    const isNotMatch = (t: string) => t.trim() !== cleanText;

    switch(type) {
      case 'spelling':
        setCorrections(prev => prev.filter(c => isNotMatch(c.wrong)));
        break;
      case 'tone':
        setToneSuggestions(prev => prev.filter(t => isNotMatch(t.current)));
        break;
      case 'style':
        setStyleSuggestions(prev => prev.filter(s => isNotMatch(s.current)));
        break;
      case 'mixing':
        setLanguageStyleMixing(prev => {
          if (!prev || !prev.corrections) return prev;
          const filtered = prev.corrections.filter(c => isNotMatch(c.current));
          return filtered.length > 0 ? { ...prev, corrections: filtered } : null;
        });
        break;
      case 'punct':
        setPunctuationIssues(prev => prev.filter(p => isNotMatch(p.currentSentence)));
        break;
      case 'euphony':
        setEuphonyImprovements(prev => prev.filter(e => isNotMatch(e.current)));
        break;
    }
  };

  const clearHighlights = async () => {
    await Word.run(async (context) => {
      context.document.body.font.highlightColor = "None";
      await context.sync();
    }).catch(console.error);
  };

  /* --- API LOGIC --- */
  const checkSpelling = async () => {
    if (!apiKey) {
      showMessage('অনুগ্রহ করে প্রথমে API Key দিন', 'error');
      setActiveModal('settings');
      return;
    }

    const text = await getTextFromWord();
    if (!text || text.trim().length === 0) {
      showMessage('টেক্সট নির্বাচন করুন বা কার্সার রাখুন', 'error');
      return;
    }

    setIsLoading(true);
    setLoadingText('বিশ্লেষণ করা হচ্ছে...');
    
    setCorrections([]);
    setToneSuggestions([]);
    setStyleSuggestions([]);
    setLanguageStyleMixing(null);
    setPunctuationIssues([]);
    setEuphonyImprovements([]);
    setContentAnalysis(null);

    await clearHighlights();

    try {
      setLoadingText('বানান ও ব্যাকরণ দেখা হচ্ছে...');
      await performMainCheck(text);

      if (selectedTone) {
        setLoadingText('টোন বিশ্লেষণ...');
        await performToneCheck(text);
      }

      if (selectedStyle !== 'none') {
        setLoadingText('ভাষারীতি বিশ্লেষণ...');
        await performStyleCheck(text);
      }

      setLoadingText('সারাংশ তৈরি হচ্ছে...');
      await analyzeContent(text);

    } catch (error) {
      console.error(error);
      showMessage('ত্রুটি হয়েছে। API Key যাচাই করুন।', 'error');
    } finally {
      setIsLoading(false);
      setLoadingText('');
    }
  };

  // IMPROVEMENT 4: Enforce JSON Mode in API Calls
  const performMainCheck = async (text: string) => {
    const response = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/${selectedModel}:generateContent?key=${apiKey}`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          contents: [{
            parts: [{
              text: `আপনি একজন দক্ষ বাংলা প্রুফরিডার। নিচের টেক্সটটি খুঁটিয়ে দেখুন।

টেক্সট:
"""
${text}
"""

⚠️ **কঠোর নির্দেশনাবলী (Strict Instructions):**
১. **বানান ভুল:** শুধুমাত্র নিশ্চিত ভুল বানান ধরুন (যুক্তাক্ষর, ণত্ব-ষত্ব)।
২. **বিরাম চিহ্ন ও প্যারাগ্রাফ:** 
   - টেক্সটের **লাইন ব্রেক (Newlines)** খেয়াল রাখুন।
   - আলাদা প্যারাগ্রাফকে জোর করে এক করবেন না।
   - **শিরোনাম, কবিতার লাইন, বা তালিকার আইটেম**-এর শেষে দাড়ি/কমা না থাকলে সেটাকে ভুল ধরবেন না।
   - শুধুমাত্র পূর্ণ বাক্যের শেষে যতিচিহ্ন না থাকলে সেটা ধরুন।
৩. **ভাষা মিশ্রণ:** সাধু ও চলিত রীতির মিশ্রণ আছে কিনা দেখুন।

⚠️ **JSON Output Rules:**
- **spellingErrors:** "wrong" ফিল্ডে শব্দটি হুবহু ইনপুট থেকে কপি করবেন।
- **punctuationIssues:** "currentSentence" ফিল্ডে ইনপুটের বাক্যটি হুবহু কপি করবেন (কোনো শব্দ যোগ/বियोग করবেন না)।

Response format (JSON):
{
  "spellingErrors": [
    {"wrong": "ভুল_শব্দ", "suggestions": ["সঠিক ১", "সঠিক ২"], "position": 0}
  ],
  "languageStyleMixing": {
    "detected": true/false,
    "recommendedStyle": "সাধু/চলিত",
    "reason": "সংক্ষিপ্ত কারণ",
    "corrections": [{"current": "শব্দ", "suggestion": "সংশোধন", "type": "সাধু→চলিত"}]
  },
  "punctuationIssues": [
    {"issue": "সমস্যা", "currentSentence": "ইনপুট বাক্য", "correctedSentence": "সংশোধিত বাক্য", "explanation": "ব্যাখ্যা"}
  ],
  "euphonyImprovements": [
    {"current": "শব্দ/বাক্যাংশ", "suggestions": ["বিকল্প"], "reason": "কেন এটি ভালো"}
  ]
}`
            }]
          }],
          // Force JSON response for reliability
          generationConfig: { responseMimeType: "application/json" }
        })
      }
    );

    const data = await response.json();
    if (!data.candidates || !data.candidates[0].content) {
       throw new Error("No content received");
    }
    
    const resultText = data.candidates[0].content.parts[0].text;
    
    // Parse JSON (Use regex to find the object even if there is markdown wrapper)
    const jsonMatch = resultText.match(/\{[\s\S]*\}/);

    if (jsonMatch) {
      try {
        const result = JSON.parse(jsonMatch[0]);
        setCorrections(result.spellingErrors || []);
        setLanguageStyleMixing(result.languageStyleMixing || null);
        setPunctuationIssues(result.punctuationIssues || []);
        setEuphonyImprovements(result.euphonyImprovements || []);

        const words = text.trim().split(/\s+/).length;
        const errors = (result.spellingErrors?.length || 0);
        setStats({
          totalWords: words,
          errorCount: errors,
          accuracy: words > 0 ? Math.round(((words - errors) / words) * 100) : 100
        });

        for (const err of (result.spellingErrors || [])) {
          await highlightInWord(err.wrong, '#fee2e2');
        }
      } catch (e) {
        console.error("JSON Parse Error", e);
      }
    }
  };

  const performToneCheck = async (text: string) => {
    const prompt = buildTonePrompt(text, selectedTone);
    const response = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/${selectedModel}:generateContent?key=${apiKey}`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ 
          contents: [{ parts: [{ text: `${prompt}\n\nযদি কোন পরিবর্তন প্রয়োজন না হয় তাহলে খালি array দিন।` }] }],
          generationConfig: { responseMimeType: "application/json" }
        })
      }
    );
    const data = await response.json();
    if (!data.candidates) return;
    
    const resultText = data.candidates[0].content.parts[0].text;
    const jsonMatch = resultText.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      const result = JSON.parse(jsonMatch[0]);
      setToneSuggestions(result.toneConversions || []);
      for (const t of (result.toneConversions || [])) {
        await highlightInWord(t.current, '#fef3c7');
      }
    }
  };

  const performStyleCheck = async (text: string) => {
    const prompt = buildStylePrompt(text, selectedStyle);
    const response = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/${selectedModel}:generateContent?key=${apiKey}`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ 
          contents: [{ parts: [{ text: `${prompt}\n\nযদি কোন পরিবর্তন প্রয়োজন না হয় তাহলে খালি array দিন।` }] }],
          generationConfig: { responseMimeType: "application/json" }
        })
      }
    );
    const data = await response.json();
    if (!data.candidates) return;

    const resultText = data.candidates[0].content.parts[0].text;
    const jsonMatch = resultText.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      const result = JSON.parse(jsonMatch[0]);
      setStyleSuggestions(result.styleConversions || []);
      for (const s of (result.styleConversions || [])) {
        await highlightInWord(s.current, '#ccfbf1');
      }
    }
  };

  const analyzeContent = async (text: string) => {
    const response = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/${selectedModel}:generateContent?key=${apiKey}`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          contents: [{
            parts: [{
              text: `বাংলা লেখাটি খুব সংক্ষেপে বিশ্লেষণ করুন:
"${text}"

Response format (JSON):
{
  "contentType": "লেখার ধরন (১-২ শব্দ)",
  "description": "খুব সংক্ষিপ্ত বর্ণনা (১ লাইন)",
  "missingElements": ["গুরুত্বপূর্ণ ১-২টি জিনিস যা নেই"],
  "suggestions": ["১টি প্রধান পরামর্শ"]
}`
            }]
          }],
          generationConfig: { responseMimeType: "application/json" }
        })
      }
    );
    const data = await response.json();
    if (!data.candidates) return;

    const resultText = data.candidates[0].content.parts[0].text;
    const jsonMatch = resultText.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      setContentAnalysis(JSON.parse(jsonMatch[0]));
    }
  };

  /* --- RENDER HELPERS --- */
  const getToneName = (t: string) => {
    const map: Record<string, string> = {
      'formal': '📋 আনুষ্ঠানিক', 'informal': '💬 অনানুষ্ঠানিক', 'professional': '💼 পেশাদার',
      'friendly': '😊 বন্ধুত্বপূর্ণ', 'respectful': '🙏 সম্মানজনক', 'persuasive': '💪 প্রভাবশালী',
      'neutral': '⚖️ নিরপেক্ষ', 'academic': '📚 শিক্ষামূলক'
    };
    return map[t] || t;
  };

  /* --- UI RENDER --- */
  return (
    <div className="app-container">
      {/* Header & Toolbar */}
      <div className="header-section">
        <div className="header-top">
          <button className="icon-btn-small" onClick={() => setActiveModal('instructions')} title="সাহায্য">❓</button>
          <div className="app-title">
            <h1>🌟 ভাষা মিত্র</h1>
            <p>বাংলা বানান ও ব্যাকরণ পরীক্ষক</p>
          </div>
          <button className="icon-btn-small" onClick={() => setActiveModal('settings')} title="সেটিংস">⚙️</button>
        </div>

        <div className="toolbar">
          <button className={`icon-btn ${selectedTone ? 'active' : ''}`} onClick={() => setActiveModal('tone')} title="টোন/ভাব নির্বাচন">
            <span className="icon">🗣️</span>
            <span className="label">টোন</span>
            {selectedTone && <span className="badge">✓</span>}
          </button>
          <button className={`icon-btn ${selectedStyle !== 'none' ? 'active' : ''}`} onClick={() => setActiveModal('style')} title="ভাষারীতি নির্বাচন">
             <span className="icon">📝</span>
            <span className="label">ভাষারীতি</span>
            {selectedStyle !== 'none' && <span className="badge">✓</span>}
          </button>
          <div style={{flex: 1}}></div>
          <button 
            onClick={checkSpelling} 
            disabled={isLoading}
            className="btn-check"
          >
            {isLoading ? '...' : '🔍 পরীক্ষা করুন'}
          </button>
        </div>
      </div>

      {/* Selection Display */}
      {(selectedTone || selectedStyle !== 'none') && (
        <div className="selection-display">
          {selectedTone && (
             <span className="selection-tag tone-tag">
               {getToneName(selectedTone)}
               <button onClick={() => setSelectedTone('')} className="clear-btn">✕</button>
             </span>
          )}
          {selectedStyle !== 'none' && (
             <span className="selection-tag style-tag">
               {selectedStyle === 'sadhu' ? '📜 সাধু রীতি' : '💬 চলিত রীতি'}
               <button onClick={() => setSelectedStyle('none')} className="clear-btn">✕</button>
             </span>
          )}
        </div>
      )}

      {/* Main Content */}
      <div className="content-area">
        
        {isLoading && (
          <div className="loading-box">
            <div className="loader"></div>
            <p>{loadingText}</p>
          </div>
        )}

        {message && (
          <div className={`message-box ${message.type}`}>
            {message.text}
          </div>
        )}

        {/* Empty State */}
        {!isLoading && stats.totalWords === 0 && !message && (
          <div className="empty-state">
            <div style={{fontSize: '40px', marginBottom: '12px'}}>✨</div>
            <p style={{fontSize: '13px', fontWeight: 500}}>সাজেশন এখানে দেখা যাবে</p>
            <p style={{fontSize: '11px', marginTop: '6px'}}>টেক্সট সিলেক্ট করে "পরীক্ষা করুন" ক্লিক করুন</p>
          </div>
        )}

        {/* Stats */}
        {stats.totalWords > 0 && (
          <div className="stats-grid">
            <div className="stat-card">
              <div className="val" style={{color: '#667eea'}}>{stats.totalWords}</div>
              <div className="lbl">শব্দ</div>
            </div>
            <div className="stat-card">
              <div className="val" style={{color: '#dc2626'}}>{stats.errorCount}</div>
              <div className="lbl">ভুল</div>
            </div>
            <div className="stat-card">
              <div className="val" style={{color: '#16a34a'}}>{stats.accuracy}%</div>
              <div className="lbl">শুদ্ধতা</div>
            </div>
          </div>
        )}

        {/* Content Analysis */}
        {contentAnalysis && (
          <>
            <div className="analysis-card content-analysis">
              <h3>📋 {contentAnalysis.contentType}</h3>
              {contentAnalysis.description && <p>{contentAnalysis.description}</p>}
            </div>
            {contentAnalysis.missingElements && contentAnalysis.missingElements.length > 0 && (
              <div className="analysis-card missing-analysis">
                <h3 style={{color:'#78350f'}}>⚠️ যা যোগ করুন</h3>
                <ul>{contentAnalysis.missingElements.map((e, i) => <li key={i}>{e}</li>)}</ul>
              </div>
            )}
             {contentAnalysis.suggestions && contentAnalysis.suggestions.length > 0 && (
              <div className="analysis-card suggestion-analysis">
                <h3 style={{color:'#115e59'}}>✨ পরামর্শ</h3>
                <ul>{contentAnalysis.suggestions.map((e, i) => <li key={i}>{e}</li>)}</ul>
              </div>
            )}
          </>
        )}

        {/* Spelling Errors */}
        {corrections.length > 0 && (
          <>
            <div className="section-header">
              <h3>📝 বানান ভুল</h3>
              <span className="section-badge" style={{background: '#fee2e2', color: '#dc2626'}}>{corrections.length}টি</span>
            </div>
            {corrections.map((c, i) => (
              <div key={i} className="suggestion-card error-card" style={{position:'relative'}} onMouseEnter={() => highlightInWord(c.wrong, '#fee2e2')}>
                <button onClick={() => dismissSuggestion('spelling', c.wrong)} className="dismiss-btn" title="বাদ দিন">✕</button>
                <div className="wrong-word">❌ {c.wrong}</div>
                {c.suggestions.map((s, j) => (
                  <button key={j} onClick={() => replaceInWord(c.wrong, s)} className="suggestion-btn success-btn">
                    ✓ {s}
                  </button>
                ))}
              </div>
            ))}
          </>
        )}

        {/* Tone Suggestions */}
        {toneSuggestions.length > 0 && (
          <>
            <div className="section-header">
              <h3>💬 টোন রূপান্তর</h3>
               <span className="section-badge" style={{background: '#fef3c7', color: '#92400e'}}>{getToneName(selectedTone)}</span>
            </div>
            {toneSuggestions.map((t, i) => (
              <div key={i} className="suggestion-card warning-card" style={{position:'relative'}} onMouseEnter={() => highlightInWord(t.current, '#fef3c7')}>
                <button onClick={() => dismissSuggestion('tone', t.current)} className="dismiss-btn" title="বাদ দিন">✕</button>
                <div className="wrong-word" style={{color: '#b45309'}}>💡 {t.current}</div>
                {t.reason && <div className="reason">{t.reason}</div>}
                <button onClick={() => replaceInWord(t.current, t.suggestion)} className="suggestion-btn warning-btn">
                  ✨ {t.suggestion}
                </button>
              </div>
            ))}
          </>
        )}

        {/* Style Suggestions */}
        {styleSuggestions.length > 0 && (
          <>
            <div className="section-header">
              <h3>📝 ভাষারীতি</h3>
               <span className="section-badge" style={{background: selectedStyle === 'sadhu' ? '#fef3c7' : '#ccfbf1', color: selectedStyle === 'sadhu' ? '#92400e' : '#0f766e'}}>
                 {selectedStyle === 'sadhu' ? '📜 সাধু রীতি' : '💬 চলিত রীতি'}
               </span>
            </div>
            {styleSuggestions.map((s, i) => (
              <div key={i} className="suggestion-card info-card" style={{borderColor: selectedStyle === 'sadhu' ? '#fbbf24' : '#5eead4', position:'relative'}} onMouseEnter={() => highlightInWord(s.current, '#ccfbf1')}>
                <button onClick={() => dismissSuggestion('style', s.current)} className="dismiss-btn" title="বাদ দিন">✕</button>
                <div style={{display:'flex', gap:'6px', alignItems:'center', marginBottom:'4px'}}>
                    <span style={{fontSize:'13px', fontWeight:600, color: selectedStyle === 'sadhu' ? '#92400e' : '#0f766e'}}>🔄 {s.current}</span>
                    {s.type && <span style={{fontSize:'9px', background:'white', padding:'2px 6px', borderRadius:'10px'}}>{s.type}</span>}
                </div>
                <button onClick={() => replaceInWord(s.current, s.suggestion)} className="suggestion-btn" style={{
                    background: selectedStyle === 'sadhu' ? 'linear-gradient(135deg, #fef3c7 0%, #fde68a 100%)' : 'linear-gradient(135deg, #ccfbf1 0%, #99f6e4 100%)',
                    borderColor: selectedStyle === 'sadhu' ? '#fbbf24' : '#5eead4',
                    color: selectedStyle === 'sadhu' ? '#92400e' : '#0f766e'
                }}>
                  ➜ {s.suggestion}
                </button>
              </div>
            ))}
          </>
        )}

        {/* Auto Style Mixing Detection */}
        {languageStyleMixing?.detected && selectedStyle === 'none' && (
          <>
            <div className="section-header">
              <h3>🔄 মিশ্রণ সনাক্ত</h3>
              <span className="section-badge" style={{background: '#e9d5ff', color: '#6b21a8'}}>স্বয়ংক্রিয়</span>
            </div>
            <div className="suggestion-card purple-card" style={{background: 'rgba(237, 233, 254, 0.5)'}}>
              <div style={{fontSize: '13px', fontWeight: 600, color: '#6b21a8'}}>
                প্রস্তাবিত: {languageStyleMixing.recommendedStyle}
              </div>
              <div style={{fontSize: '10px', color: '#6b7280', marginTop: '4px'}}>{languageStyleMixing.reason}</div>
            </div>
            {languageStyleMixing.corrections?.map((c, i) => (
              <div key={i} className="suggestion-card purple-card-light" style={{position:'relative'}} onMouseEnter={() => highlightInWord(c.current, '#e9d5ff')}>
                 <button onClick={() => dismissSuggestion('mixing', c.current)} className="dismiss-btn" title="বাদ দিন">✕</button>
                 <div style={{display:'flex', gap:'6px', alignItems:'center', marginBottom:'4px'}}>
                    <span style={{fontSize:'13px', fontWeight:600, color: '#7c3aed'}}>🔄 {c.current}</span>
                    <span style={{fontSize:'9px', background:'#e9d5ff', color:'#6b21a8', padding:'2px 6px', borderRadius:'10px'}}>{c.type}</span>
                </div>
                <button onClick={() => replaceInWord(c.current, c.suggestion)} className="suggestion-btn purple-btn">
                  ➜ {c.suggestion}
                </button>
              </div>
            ))}
          </>
        )}

        {/* Punctuation */}
        {punctuationIssues.length > 0 && (
          <>
            <div className="section-header">
               <h3>🔤 বিরাম চিহ্ন</h3>
               <span className="section-badge" style={{background: '#fed7aa', color: '#c2410c'}}>{punctuationIssues.length}টি</span>
            </div>
            {punctuationIssues.map((p, i) => (
              <div key={i} className="suggestion-card orange-card" style={{position:'relative'}} onMouseEnter={() => highlightInWord(p.currentSentence, '#ffedd5')}>
                <button onClick={() => dismissSuggestion('punct', p.currentSentence)} className="dismiss-btn" title="বাদ দিন">✕</button>
                <div className="wrong-word" style={{color: '#ea580c'}}>⚠️ {p.issue}</div>
                <div className="reason">{p.explanation}</div>
                <button onClick={() => replaceInWord(p.currentSentence, p.correctedSentence)} className="suggestion-btn orange-btn">
                  ✓ {p.correctedSentence}
                </button>
              </div>
            ))}
          </>
        )}
        
         {/* Euphony */}
        {euphonyImprovements.length > 0 && (
          <>
            <div className="section-header">
              <h3>🎵 শ্রুতিমধুরতা</h3>
               <span className="section-badge" style={{background: '#fce7f3', color: '#be185d'}}>{euphonyImprovements.length}টি</span>
            </div>
            {euphonyImprovements.map((e, i) => (
              <div key={i} className="suggestion-card" style={{borderLeft:'4px solid #db2777', position:'relative'}} onMouseEnter={() => highlightInWord(e.current, '#fce7f3')}>
                 <button onClick={() => dismissSuggestion('euphony', e.current)} className="dismiss-btn" title="বাদ দিন">✕</button>
                 <div className="wrong-word" style={{color: '#db2777'}}>🎵 {e.current}</div>
                <div className="reason">{e.reason}</div>
                {e.suggestions.map((s, j) => (
                     <button key={j} onClick={() => replaceInWord(e.current, s)} className="suggestion-btn" style={{background: '#fce7f3', borderColor: '#f9a8d4', color: '#9f1239'}}>
                      ♪ {s}
                    </button>
                ))}
              </div>
            ))}
          </>
        )}
      </div>

      {/* Footer */}
      <div className="footer">
        <p style={{fontSize:'15px', color:'rgba(255,255,255,0.9)', fontWeight:600}}>Developed by: হিমাদ্রি বিশ্বাস</p>
        <p style={{fontSize:'12px', color:'rgba(255,255,255,0.7)'}}>☎ +880 9696 196566</p>
      </div>

      {/* --- MODALS --- */}
      
      {/* Settings Modal */}
      {activeModal === 'settings' && (
        <div className="modal-overlay" onClick={() => setActiveModal('none')}>
          <div className="modal-content" onClick={e => e.stopPropagation()}>
            <div className="modal-header settings-header">
              <h3>⚙️ সেটিংস</h3>
              <button onClick={() => setActiveModal('none')}>✕</button>
            </div>
            <div className="modal-body">
              <label>🔑 Google Gemini API Key</label>
              <input type="password" value={apiKey} onChange={e => setApiKey(e.target.value)} placeholder="আপনার API Key এখানে দিন" />
              
              <label>🤖 AI Model</label>
              <select value={selectedModel} onChange={e => setSelectedModel(e.target.value)}>
                <option value="gemini-2.0-flash">Gemini 2.0 Flash (New & Fast)</option>
                <option value="gemini-1.5-flash">Gemini 1.5 Flash (Balanced)</option>
                <option value="gemini-1.5-pro">Gemini 1.5 Pro (Best Quality)</option>
              </select>
              
              <div style={{display:'flex', gap:'10px'}}>
                  <button onClick={saveSettings} className="btn-primary-full">✓ সংরক্ষণ</button>
                  <button onClick={() => setActiveModal('none')} style={{padding:'12px 20px', background:'#f3f4f6', borderRadius:'10px', border:'none', cursor:'pointer', fontWeight:600}}>বাতিল</button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Instructions Modal */}
      {activeModal === 'instructions' && (
        <div className="modal-overlay" onClick={() => setActiveModal('none')}>
          <div className="modal-content" onClick={e => e.stopPropagation()}>
            <div className="modal-header instructions-header">
              <h3>🎯 ব্যবহার নির্দেশিকা</h3>
              <button onClick={() => setActiveModal('none')}>✕</button>
            </div>
            <div className="modal-body">
              <ol style={{paddingLeft: '18px', lineHeight: '2', fontSize: '13px'}}>
                <li style={{marginBottom:'10px'}}>⚙️ সেটিংস থেকে API Key দিন</li>
                <li style={{marginBottom:'10px'}}>✍️ বাংলা টেক্সট সিলেক্ট করুন অথবা সম্পূর্ণ ডকুমেন্ট চেক করুন</li>
                <li style={{marginBottom:'10px'}}>💬 <strong>টোন</strong> আইকনে ক্লিক করে ভাব নির্বাচন করুন (ঐচ্ছিক)</li>
                <li style={{marginBottom:'10px'}}>📝 <strong>ভাষারীতি</strong> আইকনে ক্লিক করে সাধু/চলিত নির্বাচন করুন (ঐচ্ছিক)</li>
                <li style={{marginBottom:'10px'}}>🔍 "পরীক্ষা করুন" বাটনে ক্লিক করুন</li>
                <li>✓ সাজেশনে ক্লিক করে প্রতিস্থাপন করুন বা ✕ দিয়ে বাতিল করুন</li>
              </ol>
            </div>
          </div>
        </div>
      )}

      {/* Tone Modal */}
      {activeModal === 'tone' && (
        <div className="modal-overlay" onClick={() => setActiveModal('none')}>
          <div className="modal-content" onClick={e => e.stopPropagation()}>
            <div className="modal-header tone-header">
              <h3>💬 টোন/ভাব নির্বাচন</h3>
              <button onClick={() => setActiveModal('none')}>✕</button>
            </div>
            <div className="modal-body">
              {[
                {id: '', icon: '❌', title: 'কোনটি নয়', desc: 'শুধু বানান ও ব্যাকরণ পরীক্ষা'},
                {id: 'formal', icon: '📋', title: 'আনুষ্ঠানিক (Formal)', desc: 'দাপ্তরিক চিঠি, আবেদন, প্রতিবেদন'},
                {id: 'informal', icon: '💬', title: 'অনানুষ্ঠানিক (Informal)', desc: 'ব্যক্তিগত চিঠি, ব্লগ, সোশ্যাল মিডিয়া'},
                {id: 'professional', icon: '💼', title: 'পেশাদার (Professional)', desc: 'ব্যবসায়িক যোগাযোগ, কর্পোরেট'},
                {id: 'friendly', icon: '😊', title: 'বন্ধুত্বপূর্ণ (Friendly)', desc: 'উষ্ণ, আন্তরিক যোগাযোগ'},
                {id: 'respectful', icon: '🙏', title: 'সম্মানজনক (Respectful)', desc: 'বয়োজ্যেষ্ঠ বা সম্মানিত ব্যক্তি'},
                {id: 'persuasive', icon: '💪', title: 'প্রভাবশালী (Persuasive)', desc: 'মার্কেটিং, বিক্রয়, প্রচারণা'},
                {id: 'neutral', icon: '⚖️', title: 'নিরপেক্ষ (Neutral)', desc: 'সংবাদ, তথ্যমূলক লেখা'},
                {id: 'academic', icon: '📚', title: 'শিক্ষামূলক (Academic)', desc: 'গবেষণা পত্র, প্রবন্ধ'}
              ].map(opt => (
                <div key={opt.id} className={`option-item ${selectedTone === opt.id ? 'selected' : ''}`} onClick={() => {setSelectedTone(opt.id); setActiveModal('none');}}>
                  <div className="opt-icon">{opt.icon}</div>
                  <div style={{flex:1}}>
                    <div className="opt-title">{opt.title}</div>
                    <div className="opt-desc">{opt.desc}</div>
                  </div>
                  {selectedTone === opt.id && <div className="check-mark">✓</div>}
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* Style Modal */}
      {activeModal === 'style' && (
        <div className="modal-overlay" onClick={() => setActiveModal('none')}>
          <div className="modal-content" onClick={e => e.stopPropagation()}>
            <div className="modal-header style-header">
              <h3>📝 ভাষারীতি নির্বাচন</h3>
              <button onClick={() => setActiveModal('none')}>✕</button>
            </div>
            <div className="modal-body">
              {[
                {id: 'none', icon: '❌', title: 'কোনটি নয়', desc: 'স্বয়ংক্রিয় মিশ্রণ সনাক্তকরণ চালু থাকবে'},
                {id: 'sadhu', icon: '📜', title: 'সাধু রীতি', desc: 'করিতেছি, করিয়াছি, তাহার, যাহা'},
                {id: 'cholito', icon: '💬', title: 'চলিত রীতি', desc: 'করছি, করেছি, তার, যা'}
              ].map(opt => (
                <div key={opt.id} className={`option-item ${selectedStyle === opt.id ? 'selected' : ''}`} onClick={() => {setSelectedStyle(opt.id); setActiveModal('none');}}>
                  <div className="opt-icon">{opt.icon}</div>
                  <div style={{flex:1}}>
                    <div className="opt-title">{opt.title}</div>
                    <div className="opt-desc">{opt.desc}</div>
                  </div>
                  {selectedStyle === opt.id && <div className="check-mark">✓</div>}
                </div>
              ))}
              
               <div style={{padding: '10px', background: 'linear-gradient(135deg, #ede9fe 0%, #ddd6fe 100%)', borderRadius: '10px', border: '2px solid #c4b5fd', marginTop: '10px'}}>
                <h4 style={{fontSize: '12px', fontWeight: 'bold', color: '#5b21b6', marginBottom: '6px'}}>📖 পার্থক্য</h4>
                <div style={{display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '10px', fontSize: '11px'}}>
                  <div>
                    <p style={{fontWeight: 600, color: '#7c3aed', marginBottom: '2px'}}>সাধু:</p>
                    <p style={{color: '#6b7280'}}>করিতেছি, তাহার</p>
                  </div>
                  <div>
                    <p style={{fontWeight: 600, color: '#0d9488', marginBottom: '2px'}}>চলিত:</p>
                    <p style={{color: '#6b7280'}}>করছি, তার</p>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}

// ----------------------------------------------------------------------
// INITIALIZE OFFICE & REACT ENTRY POINT
// ----------------------------------------------------------------------
Office.onReady(() => {
  const rootElement = document.getElementById('root');
  if (rootElement) {
    const root = ReactDOM.createRoot(rootElement);
    root.render(<App />);
  }
});