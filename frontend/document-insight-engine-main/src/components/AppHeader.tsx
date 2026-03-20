import { FileText, Moon, Sun } from "lucide-react";
import Logo from "./Logo";
import { useState, useEffect } from "react";
import { Button } from "./ui/button";

const AppHeader = () => {
  const [isDark, setIsDark] = useState(false);

  useEffect(() => {
    // Check initial theme
    const isDarkMode = document.documentElement.classList.contains("dark");
    setIsDark(isDarkMode);
  }, []);

  const toggleTheme = () => {
    const newDark = !isDark;
    setIsDark(newDark);
    if (newDark) {
      document.documentElement.classList.add("dark");
    } else {
      document.documentElement.classList.remove("dark");
    }
  };

  return (
    <header className="gradient-header text-[hsl(var(--header-fg))] px-8 py-6 sticky top-0 z-50">
      <div className="max-w-7xl mx-auto flex items-center justify-between gap-4">
        <div className="flex items-center gap-3 animate-premium-fade overflow-hidden">
          <div className="w-12 h-12 rounded-xl bg-primary/20 backdrop-blur-sm border border-primary/30 flex items-center justify-center flex-shrink-0">
            <FileText className="w-6 h-6 text-primary" />
          </div>
          <div className="hidden sm:block">
            <h1 className="text-2xl font-bold tracking-tight text-[hsl(var(--header-fg))] whitespace-nowrap">Insurance Form Extractor</h1>
            <p className="text-sm text-[hsl(var(--header-fg))]/70">AI-Powered PDF Processing & Data Extraction</p>
          </div>
        </div>

        <div className="flex items-center gap-2 flex-shrink-0">
          <Button
            variant="ghost"
            size="icon"
            onClick={toggleTheme}
            className="text-[hsl(var(--header-fg))]/70 hover:text-[hsl(var(--header-fg))] hover:bg-black/5 dark:hover:bg-white/10 transition-colors"
            title={isDark ? "Switch to Light Mode" : "Switch to Dark Mode"}
          >
            {isDark ? <Sun className="w-5 h-5" /> : <Moon className="w-5 h-5" />}
          </Button>

          <div className="h-6 w-[1px] bg-[hsl(var(--header-fg))]/10 mx-1" />

          <Logo className="hover:scale-105 transition-transform duration-300 pl-2" />
        </div>
      </div>
    </header>
  );
};

export default AppHeader;
