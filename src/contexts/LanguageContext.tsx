import React, { createContext, useContext, useState, useEffect, ReactNode } from 'react';
import { languages, Language, LanguageKeys } from '../languages';

interface LanguageContextType {
  language: Language;
  setLanguage: (lang: Language) => void;
  t: LanguageKeys;
}

const LanguageContext = createContext<LanguageContextType | undefined>(undefined);

interface LanguageProviderProps {
  children: ReactNode;
}

export const LanguageProvider: React.FC<LanguageProviderProps> = ({ children }) => {
  const [language, setLanguage] = useState<Language>('en');

  // Load saved language from localStorage on mount
  useEffect(() => {
    const savedLanguage = localStorage.getItem('excel-processor-language') as Language;
    if (savedLanguage && languages[savedLanguage]) {
      setLanguage(savedLanguage);
    }
  }, []);

  // Save language to localStorage when it changes
  useEffect(() => {
    localStorage.setItem('excel-processor-language', language);
  }, [language]);

  const t = languages[language];

  return (
    <LanguageContext.Provider value={{ language, setLanguage, t }}>
      {children}
    </LanguageContext.Provider>
  );
};

export const useLanguage = (): LanguageContextType => {
  const context = useContext(LanguageContext);
  if (context === undefined) {
    throw new Error('useLanguage must be used within a LanguageProvider');
  }
  return context;
};

// Helper function to format messages with placeholders
export const formatMessage = (message: string, placeholders: Record<string, string | number>): string => {
  let formattedMessage = message;
  Object.entries(placeholders).forEach(([key, value]) => {
    formattedMessage = formattedMessage.replace(`{${key}}`, String(value));
  });
  return formattedMessage;
};
