import React from 'react';
import { useLanguage } from '../contexts/LanguageContext';
import { Language } from '../languages';

const LanguageSelector: React.FC = () => {
  const { language, setLanguage, t } = useLanguage();

  const languageOptions: { code: Language; name: string; flag: string }[] = [
    { code: 'en', name: 'English', flag: '🇺🇸' },
    { code: 'hu', name: 'Magyar', flag: '🇭🇺' }
  ];

  return (
    <div className="language-selector">
      <label htmlFor="language-select" className="language-label">
        {t.language}:
      </label>
      <select
        id="language-select"
        value={language}
        onChange={(e) => setLanguage(e.target.value as Language)}
        className="language-select"
      >
        {languageOptions.map((option) => (
          <option key={option.code} value={option.code}>
            {option.flag} {option.name}
          </option>
        ))}
      </select>
    </div>
  );
};

export default LanguageSelector;
