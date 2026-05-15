import React, { useState, useEffect, useMemo } from 'react';
import { initializeApp } from 'firebase/app';
import { getFirestore, doc, setDoc, onSnapshot } from 'firebase/firestore';
import { getAuth, signInAnonymously, onAuthStateChanged } from 'firebase/auth';
import { Lock, Calendar, CheckCircle, ChevronLeft, ChevronRight, Save, Download, Loader2, LogOut, User } from 'lucide-react';

// --- Firebase Configuration ---
// Canvas ఎన్విరాన్‌మెంట్‌లో ఉన్నప్పుడు ఆటోమేటిక్ కీలను వాడుతుంది. 
// లోకల్ లేదా గిట్‌హబ్‌లో రన్ చేసేటప్పుడు కింద ఉన్న ఆబ్జెక్ట్‌లో మీ వివరాలు ఇవ్వండి.
const firebaseConfig = typeof __firebase_config !== 'undefined' 
  ? JSON.parse(__firebase_config) 
  : {
      apiKey: "YOUR_API_KEY_HERE", // ఇక్కడ మీ అసలు Firebase API Key పేస్ట్ చేయండి
      authDomain: "YOUR_PROJECT.firebaseapp.com",
      projectId: "YOUR_PROJECT_ID",
      storageBucket: "YOUR_PROJECT.appspot.com",
      messagingSenderId: "YOUR_SENDER_ID",
      appId: "YOUR_APP_ID"
    };

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);
const appId = typeof __app_id !== 'undefined' ? __app_id : "tlr-attendance-v2"; 

const NAMES = ["T.L.R", "ద్రోణ", "జయంత", "శివ", "భాస్కర్", "కృష్ణ-1", "అది", "పాలినాయుడు", "శంకర్", "సాయిలు", "కృష్ణ-2", "శ్రీరాములు", "ధర్మరావు"];
const WEEKDAYS = ["ఆది", "సోమ", "మంగళ", "బుధ", "గురు", "శుక్ర", "శని"];

const generateWeeks = () => {
  const weeks = [];
  let current = new Date('2026-10-18'); 
  const end = new Date('2027-10-24'); 
  while (current <= end) { 
    weeks.push(new Date(current)); 
    current.setDate(current.getDate() + 7); 
  }
  return weeks;
};

export default function App() {
  const [user, setUser] = useState(null);
  const [isAdmin, setIsAdmin] = useState(false);
  const [loginFields, setLoginFields] = useState({ username: '', password: '' });
  const [selectedWeekIdx, setSelectedWeekIdx] = useState(0);
  const [attendanceData, setAttendanceData] = useState({});
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  
  const weeks = useMemo(() => generateWeeks(), []);
  const currentWeekStart = weeks[selectedWeekIdx];

  // Excel library loading
  useEffect(() => {
    const script = document.createElement('script');
    script.src = "https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js";
    script.async = true;
    document.body.appendChild(script);
  }, []);

  // Auth initialization
  useEffect(() => {
    const initAuth = async () => { 
      try { 
        await signInAnonymously(auth); 
      } catch (err) { 
        console.error("Firebase Auth Error:", err.message); 
      } 
    };
    initAuth();
    return onAuthStateChanged(auth, setUser);
  }, []);

  // Data fetching
  useEffect(() => {
    if (!user) return;
    setAttendanceData({});
    setLoading(true);
    const docId = `week_${currentWeekStart.toISOString().split('T')[0]}`;
    const docRef = doc(db, 'artifacts', appId, 'public', 'data', 'attendance', docId);
    
    const unsubscribe = onSnapshot(docRef, (docSnap) => {
      if (docSnap.exists()) {
        setAttendanceData(docSnap.data().records || {});
      } else {
        setAttendanceData({});
      }
      setLoading(false);
    }, (error) => {
      console.error("Firestore error:", error);
      setLoading(false);
    });
    
    return () => unsubscribe();
  }, [user, selectedWeekIdx, currentWeekStart]);

  const handleLogin = (e) => {
    e.preventDefault();
    if (loginFields.username === 'TLR' && loginFields.password === 'TLR123') {
      setIsAdmin(true);
    } else {
      alert('Invalid Login Credentials!');
    }
  };

  const toggleAttendancePart = async (name, dayIdx, partIdx) => {
    if (!isAdmin || saving) return;
    
    const newRecords = { ...attendanceData };
    if (!newRecords[name]) newRecords[name] = {};
    if (!newRecords[name][dayIdx]) newRecords[name][dayIdx] = [false, false, false];
    
    const currentParts = [...newRecords[name][dayIdx]];
    currentParts[partIdx] = !currentParts[partIdx];
    newRecords[name][dayIdx] = currentParts;
    
    setAttendanceData(newRecords);
    setSaving(true);
    
    const docId = `week_${currentWeekStart.toISOString().split('T')[0]}`;
    const docRef = doc(db, 'artifacts', appId, 'public', 'data', 'attendance', docId);
    
    try { 
      await setDoc(docRef, { records: newRecords, updatedAt: new Date().toISOString() }); 
    } catch (err) { 
      console.error("Save error:", err); 
    } finally { 
      setSaving(false); 
    }
  };

  const downloadExcel = () => {
    if (!window.XLSX) return;
    const rows = NAMES.map(name => {
      const row = { "Names": name };
      let total = 0;
      WEEKDAYS.forEach((_, i) => {
        const val = (attendanceData[name]?.[i] || [false, false, false]).filter(p => p).length * 0.5;
        row[WEEKDAYS[i]] = val;
        total += val;
      });
      row["Weekly Total"] = total;
      return row;
    });
    const ws = window.XLSX.utils.json_to_sheet(rows);
    const wb = window.XLSX.utils.book_new();
    window.XLSX.utils.book_append_sheet(wb, ws, "Attendance");
    window.XLSX.writeFile(wb, `TLR_Attendance_Week_${selectedWeekIdx + 1}.xlsx`);
  };

  if (!isAdmin) {
    return (
      <div className="min-h-screen bg-slate-50 flex items-center justify-center p-4">
        <form onSubmit={handleLogin} className="bg-white p-8 rounded-2xl shadow-lg w-full max-w-sm border border-slate-100">
          <div className="flex justify-center mb-4">
            <div className="bg-orange-100 p-3 rounded-full text-orange-600">
              <User size={32} />
            </div>
          </div>
          <h2 className="text-2xl font-bold mb-6 text-center text-slate-800">TLR Admin Login</h2>
          <div className="space-y-4">
            <input 
              type="text" 
              placeholder="Username (TLR)" 
              className="w-full p-3 border rounded-xl outline-none focus:ring-2 focus:ring-orange-500" 
              value={loginFields.username}
              onChange={e => setLoginFields({...loginFields, username: e.target.value})} 
            />
            <input 
              type="password" 
              placeholder="Password (TLR123)" 
              className="w-full p-3 border rounded-xl outline-none focus:ring-2 focus:ring-orange-500" 
              value={loginFields.password}
              onChange={e => setLoginFields({...loginFields, password: e.target.value})} 
            />
            <button className="w-full bg-orange-600 hover:bg-orange-700 text-white p-4 rounded-xl font-bold transition-colors">
              Login
            </button>
          </div>
        </form>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col font-sans">
      <header className="p-4 border-b bg-white flex flex-col md:flex-row justify-between items-center sticky top-0 z-50 shadow-sm gap-4">
        <div className="flex items-center gap-3">
          <Calendar className="text-orange-600" />
          <h1 className="text-xl font-black text-slate-800 uppercase tracking-tight">TLR Tracker</h1>
        </div>
        <div className="flex items-center gap-2">
           <button onClick={() => setSelectedWeekIdx(Math.max(0, selectedWeekIdx-1))} className="p-2 border rounded-lg hover:bg-slate-50"><ChevronLeft size={20}/></button>
           <span className="font-bold px-4 text-sm text-slate-600 bg-slate-100 py-2 rounded-lg">Week {selectedWeekIdx + 1}</span>
           <button onClick={() => setSelectedWeekIdx(Math.min(weeks.length-1, selectedWeekIdx+1))} className="p-2 border rounded-lg hover:bg-slate-50"><ChevronRight size={20}/></button>
           <div className="h-8 w-px bg-slate-200 mx-2 hidden md:block"></div>
           <button onClick={downloadExcel} className="bg-green-600 hover:bg-green-700 text-white px-4 py-2 rounded-lg font-bold flex items-center gap-2 text-sm">
             <Download size={16}/> Excel
           </button>
           <button onClick={() => setIsAdmin(false)} className="bg-slate-800 hover:bg-slate-900 text-white px-4 py-2 rounded-lg font-bold flex items-center gap-2 text-sm">
             <LogOut size={16}/> Logout
           </button>
        </div>
      </header>

      <main className="p-4 md:p-8 flex-1">
        <div className="max-w-7xl mx-auto bg-white rounded-3xl shadow-xl overflow-hidden border border-slate-200">
          {loading ? (
            <div className="p-20 flex flex-col items-center justify-center text-slate-400">
              <Loader2 className="animate-spin mb-4" size={40} />
              <p>Loading records...</p>
            </div>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full border-collapse">
                <thead>
                  <tr className="bg-slate-50 border-b border-slate-200">
                    <th className="p-4 text-left font-bold text-slate-600 sticky left-0 bg-slate-50 z-10 border-r">Names</th>
                    {WEEKDAYS.map(d => (
                      <th key={d} className="p-4 text-center font-bold text-slate-600 border-r last:border-r-0 min-w-[120px]">{d}</th>
                    ))}
                    <th className="p-4 text-center font-bold text-orange-600 bg-orange-50">Total</th>
                  </tr>
                </thead>
                <tbody>
                  {NAMES.map(name => {
                    let total = 0;
                    return (
                      <tr key={name} className="border-b border-slate-100 last:border-0 hover:bg-slate-50 transition-colors">
                        <td className="p-4 font-bold text-slate-700 sticky left-0 bg-white z-10 border-r">{name}</td>
                        {WEEKDAYS.map((_, i) => {
                          const pts = attendanceData[name]?.[i] || [false, false, false];
                          total += pts.filter(p => p).length * 0.5;
                          return (
                            <td key={i} className="p-2 border-r last:border-r-0">
                              <div className="flex gap-1 justify-center">
                                {pts.map((p, pi) => (
                                  <button 
                                    key={pi} 
                                    onClick={() => toggleAttendancePart(name, i, pi)} 
                                    disabled={saving}
                                    className={`w-7 h-7 rounded-md border flex items-center justify-center transition-all ${p ? 'bg-green-500 border-green-600 text-white shadow-sm' : 'bg-white border-slate-200 text-transparent hover:border-slate-400'}`}
                                  >
                                    <CheckCircle size={14} className={p ? 'opacity-100' : 'opacity-0'} />
                                  </button>
                                ))}
                              </div>
                            </td>
                          );
                        })}
                        <td className="p-4 text-center font-black text-orange-600 bg-orange-50">{total}</td>
                      </tr>
                    )
                  })}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </main>
      
      {saving && (
        <div className="fixed bottom-8 right-8 bg-slate-900 text-white px-6 py-3 rounded-2xl shadow-2xl flex items-center gap-3 animate-bounce z-50">
          <Loader2 className="animate-spin" size={18} />
          <span className="font-bold text-sm">Saving to Cloud...</span>
        </div>
      )}
    </div>
  );
}