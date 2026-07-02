import React from "react";
import appLogo from "@/assets/images/regenerated_image_1783006196550.png";
import webLogo from "@/assets/images/regenerated_image_1782997666816.jpg";

interface PuppyLogoProps {
  className?: string;
  size?: number;
  type?: "web" | "table";
}

export const PuppyLogo: React.FC<PuppyLogoProps> = ({ className = "", size = 44, type = "table" }) => {
  const logoSrc = type === "web" ? webLogo : appLogo;

  return (
    <div
      className={`flex items-center justify-center relative select-none group ${className}`}
      style={{ width: size, height: size }}
    >
      {/* Outer subtle glow effect */}
      <div className="absolute inset-0 bg-gradient-to-tr from-[#E5A8A0]/20 to-[#E1F1F8]/30 rounded-full blur-sm opacity-0 group-hover:opacity-100 transition-opacity duration-500" />
      
      {/* Main Container */}
      <div className="w-full h-full rounded-full bg-[#FAF5EE] border border-[#E5A8A0]/25 shadow-sm flex items-center justify-center relative overflow-hidden transition-all duration-300 group-hover:border-[#E5A8A0]/60 group-hover:shadow-md">
        <img 
          src={logoSrc} 
          alt={type === "web" ? "Web Logo" : "Table Logo"} 
          className="w-full h-full object-cover rounded-full transition-transform duration-500 group-hover:scale-105" 
          style={{ 
            imageRendering: '-webkit-optimize-contrast',
          }}
          referrerPolicy="no-referrer"
        />
      </div>
    </div>
  );
};
