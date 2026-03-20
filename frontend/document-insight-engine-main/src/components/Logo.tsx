import React from 'react';

interface LogoProps {
    className?: string;
    showText?: boolean;
    variant?: 'light' | 'dark' | 'auto';
}

const Logo: React.FC<LogoProps> = ({ className = "" }) => {
    return (
        <div className={`flex items-center justify-center ${className}`}>
            <div className="h-8 sm:h-10 flex items-center">
                <img
                    src="/Logo.png"
                    alt="CogNet Logo"
                    className="h-full w-auto object-contain select-none"
                    style={{ minWidth: '100px' }}
                />
            </div>
        </div>
    );
};

export default Logo;
