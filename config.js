// כתובת ה-Web App של Google Apps Script שכותב את ההזמנות לגיליון (מסתיימת ב-/exec).
// כל עוד השורה ריקה, האתר מציג הודעת "בהקמה" ולא שולח הזמנות.
window.SHIRTS_API_URL = 'https://script.google.com/macros/s/AKfycbzXIKwl4tuq5qU5VNWqmFRKa-afeh6EisyhKGtPMflx3RK4SroRtwm99Wxz4JnF59zZXw/exec';

// מועד סגירת ההזמנות לבאנר הטיימר בראש העמוד (תצוגה בלבד, ההזמנות לא נחסמות אחריו).
// אחרי המועד הבאנר נעלם. כדי להסתיר את הבאנר מיד, משאירים ריק: ''
window.SHIRTS_CLOSE_AT = '2026-11-22T13:00:00+02:00';
