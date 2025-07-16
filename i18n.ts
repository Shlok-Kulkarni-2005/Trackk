import i18n from 'i18next';
import { initReactI18next } from 'react-i18next';

import en from './public/locales/en/translation.json';
import sv from './public/locales/sv/translation.json';

i18n
  .use(initReactI18next)
  .init({
    resources: {
      en: { translation: en },
      sv: { translation: sv },
    },
    lng: 'en',
    fallbackLng: 'en',
    interpolation: {
      escapeValue: false,
    },
  });

export default i18n; 