// TODO improvement : one cache per language (+ default) / and or per module

const TRANSLATION_LINK = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQrSBm5JoUpSxMqNBKz5Zaud6QeXj_n1p0Ncc5tyeK15qjTrhUOpuvjBQ-5vxd_sCrsFbjOEXd5uWCc/pub?gid=1935670543&single=true&output=csv";
const TRANSLATION_CACHE_KEY = "translations"

const LANGUAGE = "fr";
const APP_LANGUAGE = "fr"
const DEFAULT_LANGUAGE = "en";

function loadTranslations() {
  const csv = UrlFetchApp.fetch(TRANSLATION_LINK);
  const data = Utilities.parseCsv(csv)
  const languages = data[0].slice(1,data.length);

  // make i18next like dictionnary
  const translations = Object.assign({}, 
    ...data.slice(1,data.length)
      .map((r) => ({
        [r[0]]: 
          Object.assign({},...languages.map((l,il) => ({[l]: r[il+1]})))})));

  const sc = CacheService.getScriptCache();
  sc.put(TRANSLATION_CACHE_KEY,Utilities.jsonStringify(translations));
}

function t(key,lang){

  const sc = CacheService.getScriptCache();
  const translations = Utilities.jsonParse(sc.get(TRANSLATION_CACHE_KEY))

  if(!lang){
    lang = key.startsWith("app") ? APP_LANGUAGE : LANGUAGE;
  }

  const r = translations[key];

  if(!r) return key;

  const tr = r[lang];

  if(!tr && lang != DEFAULT_LANGUAGE)
    return t(key, DEFAULT_LANGUAGE);     

  return tr
}