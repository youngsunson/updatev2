import { useState, useEffect } from 'react';
import ReactDOM from 'react-dom/client';
import './index.css';

/* -------------------------------------------------------------------------- */
/*                                TYPES                                       */
/* -------------------------------------------------------------------------- */

interface Correction {
  wrong: string;
  suggestions: string[];
  position?: number; // global character index in full document text
}

interface ToneSuggestion {
  current: string;
  suggestion: string;
  reason: string;
  position?: number;
}

interface StyleSuggestion {
  current: string;
  suggestion: string;
  type: string;
  position?: number;
}

interface StyleMixing {
  detected: boolean;
  recommendedStyle?: string;
  reason?: string;
  corrections?: Array<{
    current: string;
    suggestion: string;
    type: string;
    position?: number;
  }>;
}

interface PunctuationIssue {
  issue: string;
  currentSentence: string;
  correctedSentence: string;
  explanation: string;
  position?: number;
}

interface EuphonyImprovement {
  current: string;
  suggestions: string[];
  reason: string;
  position?: number;
}

interface ContentAnalysis {
  contentType: string;
  description?: string;
  missingElements?: string[];
  suggestions?: string[];
}

/* ডকুমেন্ট টাইপ এবং ভিউ ফিল্টার */
type DocumentType = 'general' | 'academic' | 'official' | 'marketing' | 'social';
type ViewFilter = 'all' | 'spelling' | 'punctuation';

const DOC_TYPE_LABELS: Record<DocumentType, string> = {
  general: 'সাধারণ লেখা',
  academic: 'একাডেমিক লেখা',
  official: 'অফিসিয়াল চিঠি',
  marketing: 'মার্কেটিং কপি',
  social: 'সোশ্যাল মিডিয়া পোস্ট'
};

/* -------------------------------------------------------------------------- */
/*                        PROMPT BUILDERS                                     */
/* -------------------------------------------------------------------------- */

const getDocTypeInstructions = (docType: DocumentType) => {
  switch (docType) {
    case 'academic':
      return `লেখাটি সম্ভবত একাডেমিক/শিক্ষামূলক। 
- যুক্তি, রেফারেন্স ও নিরপেক্ষ টোন বজায় রাখুন।
- অপ্রয়োজনীয় আবেগপ্রবণ শব্দ এড়িয়ে চলুন।
- ভাষা স্পষ্ট ও প্রমাণভিত্তিক হবে।`;
    case 'official':
      return `লেখাটি সম্ভবত অফিসিয়াল/দাপ্তরিক চিঠি।
- আনুষ্ঠানিক, সম্মানজনক ও পরিষ্কার ভাষা ব্যবহার করুন।
- অতি কথ্য/স্ল্যাং শব্দ এড়িয়ে চলুন।
- বিনীত কিন্তু দৃঢ় টোন বজায় রাখুন।`;
    case 'marketing':
      return `লেখাটি সম্ভবত মার্কেটিং/প্রচারমূলক।
- প্রভাবশালী, ইতিবাচক ও আকর্ষণীয় শব্দ ব্যবহার করুন।
- মূল সুবিধা ও প্রভাব স্পষ্ট করে তুলুন।
- অতি আনুষ্ঠানিকতা কমিয়ে সহজ, প্ররোচনামূলক ভাষা রাখুন।`;
    case 'social':
      return `লেখাটি সম্ভবত সোশ্যাল মিডিয়া/অনানুষ্ঠানিক।
- কথ্য, সহজ ও বন্ধুত্বপূর্ণ ভাষা ব্যবহার করতে পারেন।
- বেশি বড় ও জটিল বাক্য এড়িয়ে চলুন।
- পাঠকের সাথে সরাসরি কথা বলার মত টোন রাখুন।`;
    case 'general':
    default:
      return `লেখাটি সাধারণ উদ্দেশ্যের, তাই অত্যধিক আনুষ্ঠানিক বা অতি কথ্য না হয়ে মাঝামাঝি ভারসাম্য বজায় রাখুন।`;
  }
};

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
"""${text}"""

📋 **আপনার কাজ:**
1. টেক্সটের প্রতিটি শব্দ ও বাক্যাংশ বিশ্লেষণ করুন।
2. কাঙ্ক্ষিত টোনে নেই এমন শব্দগুলো চিহ্নিত করুন।
3. **গুরুত্বপূর্ণ:** "current" ফিল্ডে শব্দটি হুবহু ইনপুট টেক্সট থেকে কপি করবেন (কোনো পরিবর্তন ছাড়া)।
4. প্রতিটি সাজেশনের জন্য "position" ফিল্ডে ইনপুট টেক্সটে ওই শব্দ/ফ্রেজের শুরুর ক্যারেক্টার ইনডেক্স (০ ভিত্তিক) দিন।

📤 Response Format (ONLY valid JSON, no markdown, no extra text):
{
  "toneConversions": [
    {
      "current": "বর্তমান শব্দ (হুবহু টেক্সট থেকে)",
      "suggestion": "সংশোধিত রূপ",
      "reason": "কারণ",
      "position": 0
    }
  ]
}

যদি কোনো পরিবর্তন প্রয়োজন না হয়, তাহলে "toneConversions": [] খালি array রাখবেন।`;
};

const buildStylePrompt = (text: string, style: string) => {
  const styleInstructions: Record<string, string> = {
    'sadhu': `নিচের টেক্সটকে **সাধু রীতি**তে রূপান্তরের জন্য বিশ্লেষণ করুন। ক্রিয়াপদ (ছি->তেছি, ল->ইল), সর্বনাম (তার->তাহার) এবং অব্যয় পরিবর্তন করুন।`,
    'cholito': `নিচের টেক্সটকে **চলিত রীতি**তে রূপান্তরের জন্য বিশ্লেষণ করুন। ক্রিয়াপদ (তেছি->ছি, ইল->ল), সর্বনাম (তাহার->তার) এবং অব্যয় পরিবর্তন করুন।`
  };

  return `${styleInstructions[style]}

═══════════════════════════════════════
📝 বিশ্লেষণের জন্য টেক্সট:
"""${text}"""
═══════════════════════════════════════

⚠️ **সতর্কতা:**
- "current" ফিল্ডে শব্দটি টেক্সট থেকে **হুবহু কপি** করবেন।
- যদি কোন শব্দ পরিবর্তন প্রয়োজন না হয় তবে সেটি বাদ দিন।
- "position" ফিল্ডে ইনপুট টেক্সটে ওই শব্দের শুরুর ক্যারেক্টার ইনডেক্স (০ ভিত্তিক) দিন।

📤 Response Format (ONLY valid JSON, no markdown, no extra text):
{
  "styleConversions": [
    {
      "current": "বর্তমান শব্দ (হুবহু টেক্সট থেকে)",
      "suggestion": "সংশোধিত শব্দ",
      "type": "ক্রিয়াপদ/সর্বনাম/অব্যয়",
      "position": 0
    }
  ]
}

যদি কোনো পরিবর্তন প্রয়োজন না হয়, তাহলে "styleConversions": [] খালি array রাখবেন।`;
};

const buildMainPrompt = (text: string, docType: DocumentType) => `
আপনি একজন দক্ষ বাংলা প্রুফরিডার।

লেখার ধরন (প্রিসেট): ${DOC_TYPE_LABELS[docType]}
${getDocTypeInstructions(docType)}

নিচের টেক্সটটি খুব মনোযোগ দিয়ে বিশ্লেষণ করুন:

"""${text}"""

⚠️ কড়া নির্দেশনা:

১. বানান ভুল (spellingErrors)
   - শুধু একদম নিশ্চিত ভুল বানান ধরবেন (যেমন যুক্তাক্ষর, ণত্ব-ষত্ব, স্পষ্ট টাইপো)।
   - নাম, ব্র্যান্ড, টেকনিক্যাল টার্ম, ইংরেজি শব্দের বানান পরিবর্তন করবেন না।
   - "wrong" ফিল্ডে ইনপুটের শব্দটি হুবহু কপি করবেন।
   - "suggestions"–এ ১–৩টি বাস্তবসম্মত সঠিক বানান দিন।
   - "position": ইনপুট টেক্স্টের মধ্যে ওই ভুল শব্দের শুরুর ক্যারেক্টার ইনডেক্স (০ ভিত্তিক) দিন।

২. বিরাম চিহ্ন (punctuationIssues)
   - একমাত্র তখনই সমস্যা ধরবেন যখন:
     - পূর্ণাঙ্গ, লম্বা বাক্যের শেষে কোনো দাঁড়ি/প্রশ্নবোধক/বিস্ময়সূচক/ড্যাশ নেই।
   - শিরোনাম, তালিকা, কবিতায় দাঁড়ি না থাকলেও সেটিকে ভুল ধরবেন না।
   - "currentSentence" ফিল্ডে ইনপুট বাক্য হুবহু কপি করবেন।
   - "correctedSentence" শুধু যতিচিহ্ন/খুব সামান্য গঠন ঠিক করবে; পুরো বাক্য নতুন করে লিখবেন না।
   - "position": ইনপুট টেক্সটে ওই বাক্যের শুরুর ক্যারেক্টার ইনডেক্স (০ ভিত্তিক) দিন।

৩. ভাষারীতি মিশ্রণ (languageStyleMixing)
   - সাধু ও চলিত রীতি একসাথে ব্যবহৃত হলে তবেই detected = true করুন।
   - খুব বেশি পরিবর্তন না করে, একই ধরণের শব্দে সামঞ্জস্য আনার সাজেশন দিবেন।
   - "current" ফিল্ডে ইনপুটের অংশ হুবহু কপি করবেন।
   - প্রতিটি correction-এর জন্য "position": ইনপুট টেক্সটে ওই অংশের শুরুর ক্যারেক্টার ইনডেক্স (০ ভিত্তিক) দিন।

৪. শ্রুতিমধুরতা (euphonyImprovements)
   - কেবল তখনই সাজেশন দেবেন যখন কোনো শব্দ/বাক্যাংশ সত্যিই কানে বিরক্তিকর বা অতিরিক্ত ভারী শোনায়।
   - অর্থের বড় পরিবর্তন করবেন না, শুধু সামান্য শব্দ বাছাই ভালো করবেন।
   - "position": ইনপুট টেক্সটে ওই অংশের শুরুর ক্যারেক্টার ইনডেক্স (০ ভিত্তিক) দিন।

📤 আউটপুট ফরম্যাট (ONLY valid JSON, কোনো markdown code block, ব্যাখ্যা বা অতিরিক্ত টেক্সট নয়):

{
  "spellingErrors": [
    {
      "wrong": "ভুল_শব্দ",
      "suggestions": ["সঠিক ১", "সঠিক ২"],
      "position": 0
    }
  ],
  "languageStyleMixing": {
    "detected": true,
    "recommendedStyle": "সাধু/চলিত",
    "reason": "সংক্ষিপ্ত কারণ",
    "corrections": [
      {
        "current": "শব্দ",
        "suggestion": "সংশোধন",
        "type": "সাধু→চলিত",
        "position": 0
      }
    ]
  },
  "punctuationIssues": [
    {
      "issue": "সমস্যা",
      "currentSentence": "ইনপুট বাক্য",
      "correctedSentence": "সংশোধিত বাক্য",
      "explanation": "ব্যাখ্যা",
      "position": 0
    }
  ],
  "euphonyImprovements": [
    {
      "current": "শব্দ/বাক্যাংশ",
      "suggestions": ["বিকল্প"],
      "reason": "কেন এটি ভালো",
      "position": 0
    }
  ]
}
`;

/* -------------------------------------------------------------------------- */
/*                        TEXT CHUNKING HELPERS                               */
/* -------------------------------------------------------------------------- */

interface TextChunk {
  text: string;
  start: number; // start index in analysisText
}

const CHUNK_SIZE = 2000;

const splitIntoChunks = (text: string, chunkSize = CHUNK_SIZE): TextChunk[] => {
  const chunks: TextChunk[] = [];
  const len = text.length;
  let offset = 0;

  while (offset < len) {
    let end = Math.min(offset + chunkSize, len);

    if (end < len) {
      // কাছাকাছি কোনো newline বা দাড়ি থাকলে সেখানে কেটে দাও
      const newlineIndex = text.lastIndexOf('\n', end);
      const dandaIndex = text.lastIndexOf('।', end);
      const breakIndex = Math.max(newlineIndex, dandaIndex);
      if (breakIndex > offset + chunkSize * 0.5) {
        end = breakIndex + 1;
      }
    }

    const chunkText = text.slice(offset, end);
    chunks.push({ text: chunkText, start: offset });
    offset = end;
  }

  return chunks;
};

/* -------------------------------------------------------------------------- */
/*                           MAIN COMPONENT                                   */
/* -------------------------------------------------------------------------- */

function App() {
  // Settings State
  const [apiKey, setApiKey] = useState(localStorage.getItem('gemini_api_key') || '');
  const [selectedModel, setSelectedModel] = useState(localStorage.getItem('gemini_model') || 'gemini-2.0-flash');
  const [docType, setDocType] = useState<DocumentType>(
    (localStorage.getItem('doc_type') as DocumentType) || 'general'
  );
  
  // UI State
  const [isLoading, setIsLoading] = useState(false);
  const [loadingText, setLoadingText] = useState('');
  const [message, setMessage] = useState<{ text: string; type: 'success' | 'error' } | null>(null);
  const [activeModal, setActiveModal] = useState<'none' | 'settings' | 'instructions' | 'tone' | 'style'>('none');
  
  // Selection / Filter / Collapse
  const [selectedTone, setSelectedTone] = useState('');
  const [selectedStyle, setSelectedStyle] = useState<'none' | 'sadhu' | 'cholito'>('none');
  const [viewFilter, setViewFilter] = useState<ViewFilter>('all');
  const [collapsed, setCollapsed] = useState({
    spelling: false,
    tone: false,
    style: false,
    mixing: false,
    punctuation: false,
    euphony: false
  });

  // Data State
  const [corrections, setCorrections] = useState<Correction[]>([]);
  const [toneSuggestions, setToneSuggestions] = useState<ToneSuggestion[]>([]);
  const [styleSuggestions, setStyleSuggestions] = useState<StyleSuggestion[]>([]);
  const [languageStyleMixing, setLanguageStyleMixing] = useState<StyleMixing | null>(null);
  const [punctuationIssues, setPunctuationIssues] = useState<PunctuationIssue[]>([]);
  const [euphonyImprovements, setEuphonyImprovements] = useState<EuphonyImprovement[]>([]);
  const [contentAnalysis, setContentAnalysis] = useState<ContentAnalysis | null>(null);
  
  const [stats, setStats] = useState({ totalWords: 0, errorCount: 0, accuracy: 100 });
  const [documentText, setDocumentText] = useState<string>(''); // full document text (for position mapping)

  useEffect(() => {
    // Initialize logic if needed
  }, []);

  /* --- HELPERS --- */
  const showMessage = (text: string, type: 'success' | 'error') => {
    setMessage({ text, type });
    setTimeout(() => setMessage(null), 3500);
  };

  const saveSettings = () => {
    localStorage.setItem('gemini_api_key', apiKey);
    localStorage.setItem('gemini_model', selectedModel);
    localStorage.setItem('doc_type', docType);
    showMessage('সেটিংস সংরক্ষিত হয়েছে! ✓', 'success');
    setActiveModal('none');
  };

  // Normalizes text for comparison
  const normalize = (str: string) => {
    if (!str) return '';
    return str.trim().replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ').toLowerCase();
  };

  // position থেকে কত নম্বর occurrence, তা বের করা
  const getOccurrenceIndex = (full: string, target: string, position: number): number => {
    if (!full || !target || position == null || position < 0) return 0;
    let count = 0;
    let idx = full.indexOf(target);
    while (idx !== -1 && idx < position) {
      count++;
      idx = full.indexOf(target, idx + target.length);
    }
    return count;
  };

  const toggleSection = (key: keyof typeof collapsed) => {
    setCollapsed(prev => ({ ...prev, [key]: !prev[key] }));
  };

  /* --- WORD API INTERACTION --- */

  interface TextInfo {
    analysisText: string;  // selection or full
    fullText: string;      // always full document text
    analysisOffset: number; // analysisText শুরু হয়েছে fullText-এর কোথায়
  }

  const getTextInfoFromWord = async (): Promise<TextInfo | null> => {
    return new Promise((resolve) => {
      Word.run(async (context) => {
        const selection = context.document.getSelection();
        const body = context.document.body;
        selection.load(['text', 'isEmpty']);
        body.load('text');
        await context.sync();

        const rawFull = body.text || '';
        const rawSel = selection.text || '';

        const full = rawFull.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
        const sel = rawSel.replace(/\r\n/g, '\n').replace(/\r/g, '\n');

        let analysisText: string;
        let analysisOffset = 0;

        if (!selection.isEmpty && sel.trim().length > 0) {
          analysisText = sel;
          const idx = full.indexOf(sel);
          analysisOffset = idx >= 0 ? idx : 0;
        } else {
          analysisText = full;
          analysisOffset = 0;
        }

        resolve({ analysisText, fullText: full, analysisOffset });
      }).catch((error) => {
        console.error('Error reading Word:', error);
        resolve(null);
      });
    });
  };

  const highlightInWord = async (text: string, color: string, position?: number) => {
    const cleanText = text.trim();
    if (!cleanText) return;

    const hasSpace = /\s/.test(cleanText);

    await Word.run(async (context) => {
      const results = context.document.body.search(cleanText, { 
        matchCase: false,
        matchWholeWord: !hasSpace, // এক শব্দ হলে whole word
        ignoreSpace: true 
      });
      results.load(['items', 'items/font', 'items/text']);
      await context.sync();

      if (results.items.length === 0) return;

      if (position == null || !documentText) {
        // position নাই, সবগুলো হাইলাইট
        for (let i = 0; i < results.items.length; i++) {
          results.items[i].font.highlightColor = color;
        }
      } else {
        const occurrenceIndex = getOccurrenceIndex(documentText, cleanText, position);
        const idx = Math.min(occurrenceIndex, results.items.length - 1);
        results.items[idx].font.highlightColor = color;
      }

      await context.sync();
    }).catch(console.error);
  };

  const replaceInWord = async (oldText: string, newText: string, position?: number) => {
    const cleanOldText = oldText.trim();
    if (!cleanOldText) return;
    
    const hasSpace = /\s/.test(cleanOldText);
    let success = false;

    await Word.run(async (context) => {
      const results = context.document.body.search(cleanOldText, { 
        matchCase: false,
        matchWholeWord: !hasSpace,
        ignoreSpace: true 
      });
      results.load('items');
      await context.sync();

      if (results.items.length === 0) return;

      if (position == null || !documentText) {
        // সব occurrence রিপ্লেস
        results.items.forEach((item) => {
          item.insertText(newText, Word.InsertLocation.replace);
          item.font.highlightColor = "None";
        });
      } else {
        const occurrenceIndex = getOccurrenceIndex(documentText, cleanOldText, position);
        const idx = Math.min(occurrenceIndex, results.items.length - 1);
        const item = results.items[idx];
        item.insertText(newText, Word.InsertLocation.replace);
        item.font.highlightColor = "None";
      }

      await context.sync();
      success = true;
    }).catch(console.error);

    if (success) {
      const target = normalize(cleanOldText);
      const isNotMatch = (textToCheck: string) => normalize(textToCheck) !== target;

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
      showMessage(`শব্দটি ডকুমেন্টে খুঁজে পাওয়া যায়নি।`, 'error');
    }
  };

  const dismissSuggestion = (type: 'spelling' | 'tone' | 'style' | 'mixing' | 'punct' | 'euphony', textToDismiss: string) => {
    const target = normalize(textToDismiss);
    const isNotMatch = (t: string) => normalize(t) !== target;

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

  /* --- GEMINI JSON HELPER --- */
  const callGeminiJson = async (
    prompt: string,
    { temperature = 0.2 }: { temperature?: number } = {}
  ): Promise<any | null> => {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${selectedModel}:generateContent?key=${apiKey}`;

    try {
      const response = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: {
            responseMimeType: 'application/json',
            temperature
          }
        })
      });

      if (!response.ok) {
        let userMessage = 'অজানা ত্রুটি ঘটেছে। পরে আবার চেষ্টা করুন।';
        if (response.status === 401 || response.status === 403) {
          userMessage = 'API Key বা অনুমতি সংক্রান্ত সমস্যা। সেটিংসে আপনার Key যাচাই করুন।';
        } else if (response.status === 429) {
          userMessage = 'Rate limit অতিক্রম করেছে। কিছুক্ষণ পর আবার চেষ্টা করুন।';
        } else if (response.status >= 500) {
          userMessage = 'সার্ভার-সাইড সমস্যা (5xx)। কিছুক্ষণ পর আবার চেষ্টা করুন।';
        }

        const err: any = new Error(userMessage);
        err.status = response.status;
        err.statusText = response.statusText;
        try {
          err.body = await response.text();
        } catch {
          // ignore
        }
        throw err;
      }

      const data = await response.json();

      const parts = data?.candidates?.[0]?.content?.parts;
      if (!Array.isArray(parts)) return null;

      const raw = parts.map((p: any) => p.text ?? '').join('').trim();
      if (!raw) return null;

      try {
        return JSON.parse(raw);
      } catch {
        const match = raw.match(/\{[\s\S]*\}/);
        if (match) {
          try {
            return JSON.parse(match[0]);
          } catch {
            return null;
          }
        }
        return null;
      }
    } catch (err: any) {
      if (err?.name === 'TypeError') {
        throw new Error('ইন্টারনেট সংযোগ বা নেটওয়ার্ক সমস্যার কারণে অনুরোধ সম্পন্ন হয়নি।');
      }
      throw err;
    }
  };

  /* --- MAIN CHECK (CHUNKED) --- */

  interface MainCheckChunkResult {
    spellingErrors: Correction[];
    mixing: StyleMixing | null;
    punctuation: PunctuationIssue[];
    euphony: EuphonyImprovement[];
  }

  const performMainCheckOnChunk = async (
    chunk: TextChunk,
    baseOffset: number
  ): Promise<MainCheckChunkResult> => {
    const result = await callGeminiJson(buildMainPrompt(chunk.text, docType), { temperature: 0.1 });
    if (!result) throw new Error('Gemini থেকে ফলাফল পাওয়া যায়নি।');

    const chunkBase = baseOffset + chunk.start;

    const spellingErrors: Correction[] = (result.spellingErrors || []).map((err: any) => {
      let localPos: number =
        typeof err.position === 'number'
          ? err.position
          : chunk.text.indexOf(err.wrong);
      if (localPos < 0) localPos = 0;
      const globalPos = chunkBase + localPos;
      return {
        wrong: err.wrong,
        suggestions: err.suggestions || [],
        position: globalPos
      };
    });

    const punctuation: PunctuationIssue[] = (result.punctuationIssues || []).map((p: any) => {
      let localPos: number =
        typeof p.position === 'number'
          ? p.position
          : chunk.text.indexOf(p.currentSentence);
      if (localPos < 0) localPos = 0;
      const globalPos = chunkBase + localPos;
      return {
        issue: p.issue,
        currentSentence: p.currentSentence,
        correctedSentence: p.correctedSentence,
        explanation: p.explanation,
        position: globalPos
      };
    });

    const euphony: EuphonyImprovement[] = (result.euphonyImprovements || []).map((e: any) => {
      let localPos: number =
        typeof e.position === 'number'
          ? e.position
          : chunk.text.indexOf(e.current);
      if (localPos < 0) localPos = 0;
      const globalPos = chunkBase + localPos;
      return {
        current: e.current,
        suggestions: e.suggestions || [],
        reason: e.reason,
        position: globalPos
      };
    });

    let mixing: StyleMixing | null = null;
    if (result.languageStyleMixing) {
      const m = result.languageStyleMixing;
      let corrections: StyleMixing['corrections'] = undefined;

      if (Array.isArray(m.corrections)) {
        corrections = m.corrections.map((c: any) => {
          let localPos: number =
            typeof c.position === 'number'
              ? c.position
              : chunk.text.indexOf(c.current);
          if (localPos < 0) localPos = 0;
          const globalPos = chunkBase + localPos;
          return {
            current: c.current,
            suggestion: c.suggestion,
            type: c.type,
            position: globalPos
          };
        });
      }

      mixing = {
        detected: !!m.detected,
        recommendedStyle: m.recommendedStyle,
        reason: m.reason,
        corrections
      };
    }

    return { spellingErrors, mixing, punctuation, euphony };
  };

  const runMainCheckOverChunks = async (analysisText: string, baseOffset: number) => {
    const chunks = splitIntoChunks(analysisText);
    const allCorrections: Correction[] = [];
    const allPunct: PunctuationIssue[] = [];
    const allEuphony: EuphonyImprovement[] = [];
    let mixingAggregate: StyleMixing | null = null;

    for (let i = 0; i < chunks.length; i++) {
      const chunk = chunks[i];
      setLoadingText(`বানান ও ব্যাকরণ দেখা হচ্ছে... (${i + 1}/${chunks.length})`);
      const { spellingErrors, mixing, punctuation, euphony } =
        await performMainCheckOnChunk(chunk, baseOffset);

      allCorrections.push(...spellingErrors);
      allPunct.push(...punctuation);
      allEuphony.push(...euphony);

      if (mixing?.detected) {
        if (!mixingAggregate) {
          mixingAggregate = { ...mixing, corrections: mixing.corrections || [] };
        } else {
          mixingAggregate.detected = true;
          if (!mixingAggregate.corrections) mixingAggregate.corrections = [];
          if (mixing.corrections) {
            mixingAggregate.corrections.push(...mixing.corrections);
          }
          // recommendedStyle ভিন্ন হলে এখন আপাতত প্রথমটাকেই রেখে দিচ্ছি
        }
      }
    }

    // ডকুমেন্টে যে ক্রমে আছে সেই ক্রমে sort
    allCorrections.sort((a, b) => (a.position ?? 0) - (b.position ?? 0));
    allPunct.sort((a, b) => (a.position ?? 0) - (b.position ?? 0));
    allEuphony.sort((a, b) => (a.position ?? 0) - (b.position ?? 0));

    if (mixingAggregate?.corrections) {
      mixingAggregate.corrections.sort(
        (a, b) => (a.position ?? 0) - (b.position ?? 0)
      );
    }

    setCorrections(allCorrections);
    setPunctuationIssues(allPunct);
    setEuphonyImprovements(allEuphony);
    setLanguageStyleMixing(mixingAggregate);

    const words = analysisText.trim().split(/\s+/).filter(Boolean).length;
    const errors = allCorrections.length;
    setStats({
      totalWords: words,
      errorCount: errors,
      accuracy: words > 0 ? Math.round(((words - errors) / words) * 100) : 100
    });

    // বানান ভুল হাইলাইট (position-aware)
    for (const err of allCorrections) {
      await highlightInWord(err.wrong, '#fee2e2', err.position);
    }
  };

  /* --- OTHER CHECKS --- */

  const performToneCheck = async (text: string, baseOffset: number) => {
    const prompt = buildTonePrompt(text, selectedTone);
    const result = await callGeminiJson(
      `${prompt}\n\nযদি কোন পরিবর্তন প্রয়োজন না হয় তাহলে "toneConversions": [] খালি array রাখবেন।`,
      { temperature: 0.2 }
    );
    if (!result) return;

    const toneConversions: ToneSuggestion[] = (result.toneConversions || []).map((t: any) => {
      let localPos: number =
        typeof t.position === 'number'
          ? t.position
          : text.indexOf(t.current);
      if (localPos < 0) localPos = 0;
      const globalPos = baseOffset + localPos;
      return {
        current: t.current,
        suggestion: t.suggestion,
        reason: t.reason,
        position: globalPos
      };
    });

    toneConversions.sort((a, b) => (a.position ?? 0) - (b.position ?? 0));
    setToneSuggestions(toneConversions);

    for (const t of toneConversions) {
      await highlightInWord(t.current, '#fef3c7', t.position);
    }
  };

  const performStyleCheck = async (text: string, baseOffset: number) => {
    const prompt = buildStylePrompt(text, selectedStyle);
    const result = await callGeminiJson(
      `${prompt}\n\nযদি কোন পরিবর্তন প্রয়োজন না হয় তাহলে "styleConversions": [] খালি array রাখবেন।`,
      { temperature: 0.2 }
    );
    if (!result) return;

    const styleConversions: StyleSuggestion[] = (result.styleConversions || []).map((s: any) => {
      let localPos: number =
        typeof s.position === 'number'
          ? s.position
          : text.indexOf(s.current);
      if (localPos < 0) localPos = 0;
      const globalPos = baseOffset + localPos;
      return {
        current: s.current,
        suggestion: s.suggestion,
        type: s.type,
        position: globalPos
      };
    });

    styleConversions.sort((a, b) => (a.position ?? 0) - (b.position ?? 0));
    setStyleSuggestions(styleConversions);

    for (const s of styleConversions) {
      await highlightInWord(s.current, '#ccfbf1', s.position);
    }
  };

  const analyzeContent = async (text: string, docType: DocumentType) => {
    const prompt = `
বাংলা লেখাটি খুব সংক্ষেপে বিশ্লেষণ করুন।

প্রিসেট ডকুমেন্ট টাইপ (হিন্ট): ${DOC_TYPE_LABELS[docType]}
যদি আপনার বিশ্লেষণ অনুযায়ী ধরন ভিন্ন মনে হয়, "contentType" ফিল্ডে তা সংক্ষেপে উল্লেখ করুন।

লেখা:
"""${text}"""

Response format (ONLY valid JSON, no extra text):

{
  "contentType": "লেখার ধরন (১-২ শব্দ)",
  "description": "খুব সংক্ষিপ্ত বর্ণনা (১ লাইন)",
  "missingElements": ["গুরুত্বপূর্ণ ১-২টি জিনিস যা নেই"],
  "suggestions": ["১টি প্রধান পরামর্শ"]
}
`;
    const result = await callGeminiJson(prompt, { temperature: 0.5 });
    if (!result) return;

    setContentAnalysis(result as ContentAnalysis);
  };

  /* --- MAIN ENTRY: CHECK SPELLING --- */

  const checkSpelling = async () => {
    if (!apiKey) {
      showMessage('অনুগ্রহ করে প্রথমে API Key দিন', 'error');
      setActiveModal('settings');
      return;
    }

    const info = await getTextInfoFromWord();
    if (!info || !info.analysisText || info.analysisText.trim().length === 0) {
      showMessage('টেক্সট নির্বাচন করুন বা কার্সার রাখুন', 'error');
      return;
    }

    const { analysisText, fullText, analysisOffset } = info;
    setDocumentText(fullText);

    setIsLoading(true);
    setLoadingText('বিশ্লেষণ করা হচ্ছে...');
    
    setCorrections([]);
    setToneSuggestions([]);
    setStyleSuggestions([]);
    setLanguageStyleMixing(null);
    setPunctuationIssues([]);
    setEuphonyImprovements([]);
    setContentAnalysis(null);
    setStats({ totalWords: 0, errorCount: 0, accuracy: 100 });

    await clearHighlights();

    try {
      await runMainCheckOverChunks(analysisText, analysisOffset);

      const extraTasks: Promise<void>[] = [];

      if (selectedTone) {
        extraTasks.push((async () => {
          setLoadingText('টোন বিশ্লেষণ হচ্ছে...');
          await performToneCheck(analysisText, analysisOffset);
        })());
      }

      if (selectedStyle !== 'none') {
        extraTasks.push((async () => {
          setLoadingText('ভাষারীতি বিশ্লেষণ হচ্ছে...');
          await performStyleCheck(analysisText, analysisOffset);
        })());
      }

      extraTasks.push((async () => {
        setLoadingText('সারাংশ তৈরি হচ্ছে...');
        await analyzeContent(analysisText, docType);
      })());

      await Promise.all(extraTasks);
    } catch (error: any) {
      console.error(error);
      showMessage(
        error?.message || 'ত্রুটি হয়েছে। API Key, Model বা নেটওয়ার্ক চেক করুন।',
        'error'
      );
    } finally {
      setIsLoading(false);
      setLoadingText('');
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
          <span className="selection-tag" style={{background:'#e5e7eb', color:'#111827'}}>
            📄 {DOC_TYPE_LABELS[docType]}
          </span>
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
          <>
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

            {/* Filter Bar */}
            <div className="filter-bar">
              <button
                className={`filter-btn ${viewFilter === 'all' ? 'active' : ''}`}
                onClick={() => setViewFilter('all')}
              >
                সব দেখাও
              </button>
              <button
                className={`filter-btn ${viewFilter === 'spelling' ? 'active' : ''}`}
                onClick={() => setViewFilter('spelling')}
              >
                শুধু বানান
              </button>
              <button
                className={`filter-btn ${viewFilter === 'punctuation' ? 'active' : ''}`}
                onClick={() => setViewFilter('punctuation')}
              >
                শুধু বিরামচিহ্ন
              </button>
            </div>
          </>
        )}

        {/* Content Analysis */}
        {contentAnalysis && viewFilter === 'all' && (
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
        {corrections.length > 0 && (viewFilter === 'all' || viewFilter === 'spelling') && (
          <>
            <div className="section-header">
              <div style={{display:'flex', alignItems:'center', gap:6}}>
                <button
                  className="collapse-btn"
                  onClick={() => toggleSection('spelling')}
                  title={collapsed.spelling ? 'Expand' : 'Collapse'}
                >
                  {collapsed.spelling ? '＋' : '−'}
                </button>
                <h3>📝 বানান ভুল</h3>
              </div>
              <span className="section-badge" style={{background: '#fee2e2', color: '#dc2626'}}>{corrections.length}টি</span>
            </div>
            {!collapsed.spelling && corrections.map((c, i) => (
              <div
                key={i}
                className="suggestion-card error-card"
                style={{position:'relative'}}
                onMouseEnter={() => highlightInWord(c.wrong, '#fee2e2', c.position)}
              >
                <button onClick={() => dismissSuggestion('spelling', c.wrong)} className="dismiss-btn" title="বাদ দিন">✕</button>
                <div className="wrong-word">❌ {c.wrong}</div>
                {c.suggestions.map((s, j) => (
                  <button
                    key={j}
                    onClick={() => replaceInWord(c.wrong, s, c.position)}
                    className="suggestion-btn success-btn"
                  >
                    ✓ {s}
                  </button>
                ))}
              </div>
            ))}
          </>
        )}

        {/* Tone Suggestions */}
        {toneSuggestions.length > 0 && viewFilter === 'all' && (
          <>
            <div className="section-header">
              <div style={{display:'flex', alignItems:'center', gap:6}}>
                <button
                  className="collapse-btn"
                  onClick={() => toggleSection('tone')}
                  title={collapsed.tone ? 'Expand' : 'Collapse'}
                >
                  {collapsed.tone ? '＋' : '−'}
                </button>
                <h3>💬 টোন রূপান্তর</h3>
              </div>
               <span className="section-badge" style={{background: '#fef3c7', color: '#92400e'}}>{getToneName(selectedTone)}</span>
            </div>
            {!collapsed.tone && toneSuggestions.map((t, i) => (
              <div
                key={i}
                className="suggestion-card warning-card"
                style={{position:'relative'}}
                onMouseEnter={() => highlightInWord(t.current, '#fef3c7', t.position)}
              >
                <button onClick={() => dismissSuggestion('tone', t.current)} className="dismiss-btn" title="বাদ দিন">✕</button>
                <div className="wrong-word" style={{color: '#b45309'}}>💡 {t.current}</div>
                {t.reason && <div className="reason">{t.reason}</div>}
                <button
                  onClick={() => replaceInWord(t.current, t.suggestion, t.position)}
                  className="suggestion-btn warning-btn"
                >
                  ✨ {t.suggestion}
                </button>
              </div>
            ))}
          </>
        )}

        {/* Style Suggestions */}
        {styleSuggestions.length > 0 && viewFilter === 'all' && (
          <>
            <div className="section-header">
              <div style={{display:'flex', alignItems:'center', gap:6}}>
                <button
                  className="collapse-btn"
                  onClick={() => toggleSection('style')}
                  title={collapsed.style ? 'Expand' : 'Collapse'}
                >
                  {collapsed.style ? '＋' : '−'}
                </button>
                <h3>📝 ভাষারীতি</h3>
              </div>
               <span className="section-badge" style={{background: selectedStyle === 'sadhu' ? '#fef3c7' : '#ccfbf1', color: selectedStyle === 'sadhu' ? '#92400e' : '#0f766e'}}>
                 {selectedStyle === 'sadhu' ? '📜 সাধু রীতি' : '💬 চলিত রীতি'}
               </span>
            </div>
            {!collapsed.style && styleSuggestions.map((s, i) => (
              <div
                key={i}
                className="suggestion-card info-card"
                style={{borderColor: selectedStyle === 'sadhu' ? '#fbbf24' : '#5eead4', position:'relative'}}
                onMouseEnter={() => highlightInWord(s.current, '#ccfbf1', s.position)}
              >
                <button onClick={() => dismissSuggestion('style', s.current)} className="dismiss-btn" title="বাদ দিন">✕</button>
                <div style={{display:'flex', gap:'6px', alignItems:'center', marginBottom:'4px'}}>
                    <span style={{fontSize:'13px', fontWeight:600, color: selectedStyle === 'sadhu' ? '#92400e' : '#0f766e'}}>🔄 {s.current}</span>
                    {s.type && <span style={{fontSize:'9px', background:'white', padding:'2px 6px', borderRadius:'10px'}}>{s.type}</span>}
                </div>
                <button
                  onClick={() => replaceInWord(s.current, s.suggestion, s.position)}
                  className="suggestion-btn"
                  style={{
                    background: selectedStyle === 'sadhu' ? 'linear-gradient(135deg, #fef3c7 0%, #fde68a 100%)' : 'linear-gradient(135deg, #ccfbf1 0%, #99f6e4 100%)',
                    borderColor: selectedStyle === 'sadhu' ? '#fbbf24' : '#5eead4',
                    color: selectedStyle === 'sadhu' ? '#92400e' : '#0f766e'
                  }}
                >
                  ➜ {s.suggestion}
                </button>
              </div>
            ))}
          </>
        )}

        {/* Auto Style Mixing Detection */}
        {languageStyleMixing?.detected && selectedStyle === 'none' && viewFilter === 'all' && (
          <>
            <div className="section-header">
              <div style={{display:'flex', alignItems:'center', gap:6}}>
                <button
                  className="collapse-btn"
                  onClick={() => toggleSection('mixing')}
                  title={collapsed.mixing ? 'Expand' : 'Collapse'}
                >
                  {collapsed.mixing ? '＋' : '−'}
                </button>
                <h3>🔄 মিশ্রণ সনাক্ত</h3>
              </div>
              <span className="section-badge" style={{background: '#e9d5ff', color: '#6b21a8'}}>স্বয়ংক্রিয়</span>
            </div>
            {!collapsed.mixing && (
              <>
                <div className="suggestion-card purple-card" style={{background: 'rgba(237, 233, 254, 0.5)'}}>
                  <div style={{fontSize: '13px', fontWeight: 600, color: '#6b21a8'}}>
                    প্রস্তাবিত: {languageStyleMixing.recommendedStyle}
                  </div>
                  <div style={{fontSize: '10px', color: '#6b7280', marginTop: '4px'}}>{languageStyleMixing.reason}</div>
                </div>
                {languageStyleMixing.corrections?.map((c, i) => (
                  <div
                    key={i}
                    className="suggestion-card purple-card-light"
                    style={{position:'relative'}}
                    onMouseEnter={() => highlightInWord(c.current, '#e9d5ff', c.position)}
                  >
                    <button onClick={() => dismissSuggestion('mixing', c.current)} className="dismiss-btn" title="বাদ দিন">✕</button>
                    <div style={{display:'flex', gap:'6px', alignItems:'center', marginBottom:'4px'}}>
                      <span style={{fontSize:'13px', fontWeight:600, color: '#7c3aed'}}>🔄 {c.current}</span>
                      <span style={{fontSize:'9px', background:'#e9d5ff', color:'#6b21a8', padding:'2px 6px', borderRadius:'10px'}}>{c.type}</span>
                    </div>
                    <button
                      onClick={() => replaceInWord(c.current, c.suggestion, c.position)}
                      className="suggestion-btn purple-btn"
                    >
                      ➜ {c.suggestion}
                    </button>
                  </div>
                ))}
              </>
            )}
          </>
        )}

        {/* Punctuation */}
        {punctuationIssues.length > 0 && (viewFilter === 'all' || viewFilter === 'punctuation') && (
          <>
            <div className="section-header">
              <div style={{display:'flex', alignItems:'center', gap:6}}>
                <button
                  className="collapse-btn"
                  onClick={() => toggleSection('punctuation')}
                  title={collapsed.punctuation ? 'Expand' : 'Collapse'}
                >
                  {collapsed.punctuation ? '＋' : '−'}
                </button>
                <h3>🔤 বিরাম চিহ্ন</h3>
              </div>
               <span className="section-badge" style={{background: '#fed7aa', color: '#c2410c'}}>{punctuationIssues.length}টি</span>
            </div>
            {!collapsed.punctuation && punctuationIssues.map((p, i) => (
              <div
                key={i}
                className="suggestion-card orange-card"
                style={{position:'relative'}}
                onMouseEnter={() => highlightInWord(p.currentSentence, '#ffedd5', p.position)}
              >
                <button onClick={() => dismissSuggestion('punct', p.currentSentence)} className="dismiss-btn" title="বাদ দিন">✕</button>
                <div className="wrong-word" style={{color: '#ea580c'}}>⚠️ {p.issue}</div>
                <div className="reason">{p.explanation}</div>
                <button
                  onClick={() => replaceInWord(p.currentSentence, p.correctedSentence, p.position)}
                  className="suggestion-btn orange-btn"
                >
                  ✓ {p.correctedSentence}
                </button>
              </div>
            ))}
          </>
        )}
        
         {/* Euphony */}
        {euphonyImprovements.length > 0 && shouldShowSection('euphony') && (
          <>
            <div className="section-header">
              <h3>🎵 শ্রুতিমধুরতা</h3>
               <span className="section-badge" style={{background: '#fce7f3', color: '#be185d'}}>{euphonyImprovements.length}টি</span>
               <button
                 className="collapse-btn"
                 onClick={() => toggleSection('euphony')}
               >
                 {collapsedSections.euphony ? '➕' : '➖'}
               </button>
            </div>
            {!collapsedSections.euphony && euphonyImprovements.map((e, i) => (
              <div
                key={i}
                className="suggestion-card"
                style={{borderLeft:'4px solid #db2777', position:'relative'}}
                onMouseEnter={() => highlightInWord(e.current, '#fce7f3', e.position)}
              >
                 <button onClick={() => dismissSuggestion('euphony', e.current)} className="dismiss-btn" title="বাদ দিন">✕</button>
                 <div className="wrong-word" style={{color: '#db2777'}}>🎵 {e.current}</div>
                <div className="reason">{e.reason}</div>
                {e.suggestions.map((s, j) => (
                     <button
                       key={j}
                       onClick={() => replaceInWord(e.current, s, e.position)}
                       className="suggestion-btn"
                       style={{background: '#fce7f3', borderColor: '#f9a8d4', color: '#9f1239'}}
                     >
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

              <label>📂 ডকুমেন্ট টাইপ (ডিফল্ট)</label>
              <select value={docType} onChange={e => setDocType(e.target.value as DocType)}>
                <option value="generic">সাধারণ লেখা</option>
                <option value="academic">একাডেমিক লেখা</option>
                <option value="official">অফিশিয়াল চিঠি</option>
                <option value="marketing">মার্কেটিং কপি</option>
                <option value="social">সোশ্যাল মিডিয়া পোস্ট</option>
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
                <li style={{marginBottom:'10px'}}>📂 প্রয়োজন হলে ডক টাইপ নির্বাচন করুন (একাডেমিক/অফিসিয়াল/মার্কেটিং ইত্যাদি)</li>
                <li style={{marginBottom:'10px'}}>✍️ বাংলা টেক্সট সিলেক্ট করুন অথবা সম্পূর্ণ ডকুমেন্ট চেক করুন</li>
                <li style={{marginBottom:'10px'}}>💬 <strong>টোন</strong> আইকনে ক্লিক করে ভাব নির্বাচন করুন (ঐচ্ছিক)</li>
                <li style={{marginBottom:'10px'}}>📝 <strong>ভাষারীতি</strong> আইকনে ক্লিক করে সাধু/চলিত নির্বাচন করুন (ঐচ্ছিক)</li>
                <li style={{marginBottom:'10px'}}>🔍 "পরীক্ষা করুন" বাটনে ক্লিক করুন</li>
                <li style={{marginBottom:'10px'}}>🔎 উপরের ফিল্টার থেকে "শুধু বানান / শুধু বিরামচিহ্ন / সব" বেছে নিন</li>
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
                <div
                  key={opt.id}
                  className={`option-item ${selectedTone === opt.id ? 'selected' : ''}`}
                  onClick={() => { setSelectedTone(opt.id); setActiveModal('none'); }}
                >
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
                <div
                  key={opt.id}
                  className={`option-item ${selectedStyle === opt.id ? 'selected' : ''}`}
                  onClick={() => { setSelectedStyle(opt.id as any); setActiveModal('none'); }}
                >
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

      {/* Doc Type Modal */}
      {activeModal === 'doctype' && (
        <div className="modal-overlay" onClick={() => setActiveModal('none')}>
          <div className="modal-content" onClick={e => e.stopPropagation()}>
            <div className="modal-header style-header">
              <h3>📂 ডকুমেন্ট টাইপ নির্বাচন</h3>
              <button onClick={() => setActiveModal('none')}>✕</button>
            </div>
            <div className="modal-body">
              {(['generic', 'academic', 'official', 'marketing', 'social'] as DocType[]).map(dt => {
                const cfg = DOC_TYPE_CONFIG[dt];
                return (
                  <div
                    key={dt}
                    className={`option-item ${docType === dt ? 'selected' : ''}`}
                    onClick={() => {
                      setDocType(dt);
                      // যদি tone আগে থেকে সেট না থাকে তবে ডক টাইপের default tone সেট করা
                      if (!selectedTone && cfg.defaultTone) {
                        setSelectedTone(cfg.defaultTone);
                      }
                      setActiveModal('none');
                    }}
                  >
                    <div className="opt-icon">📂</div>
                    <div style={{flex:1}}>
                      <div className="opt-title">{cfg.label}</div>
                      <div className="opt-desc">{cfg.description}</div>
                    </div>
                    {docType === dt && <div className="check-mark">✓</div>}
                  </div>
                );
              })}
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
