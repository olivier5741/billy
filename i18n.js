// TODO improvement : one cache per language (+ default) / and or per module

const TRANSLATION_LINK = "https://docs.google.com/spreadsheets/d/e/2PACX-1vR_Ju3OBcUuNHuPeEM1q1a_T2rIuZEry7nsK1Jvxol13VVzWHkB_J3uiQjUHNxDgAyRd7_oKjZPonJG/pub?gid=1935670543&single=true&output=csv";
const TRANSLATION_CACHE_KEY = "translations"

const LANGUAGE = "en";
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
  return translations;
}

function getTranslations(){
  try{
    const sc = CacheService.getScriptCache();
    const json = sc.get(TRANSLATION_CACHE_KEY);

    if(json){
      return Utilities.jsonParse(json)
    }else{
      return loadTranslations();
    }


  }catch(err){
    Logger.log(err)
  }
}

function t(key,lang){

  const translations = getTranslations();

  if(!translations)
    return key;

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

function rt(keyStartsWith,name,lang = APP_LANGUAGE){
  const translations = getTranslations();

  // TODO slow performance
  for(let [key, value] of Object.entries(translations)){
    if(key.startsWith(keyStartsWith) == false)
      continue

    if(value[lang] == name)
      return key
  }

  return name
}

function testRt(){
  Logger.log(rt("app.stock.sheet.name.","sortie","fr"))
}