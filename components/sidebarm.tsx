"use client";
import "../i18n";
import React, { useState, useEffect } from "react";
import { useRouter } from "next/navigation";
import {
  X,
  User,
  Globe,
  Settings,
  Clock,
  LogOut,
  ArrowLeft,
  Languages,
} from "lucide-react";
import { useTranslation } from "react-i18next";

// Type definitions
interface DevicePreferenceContentProps {
  onBack: () => void;
  appParams: {
    location: string;
    serverReachability: string;
    dataConnections: string;
    latency: string;
  };
  deviceParams: {
    platform: string;
    platformDetails: string;
    manufacturer: string;
    serviceProvider: string;
    serviceProviderDetails: string;
  };
}

interface SettingsContentProps {
  onBack: () => void;
}

interface MainMenuContentProps {
  username: string | null;
  onDevicePreference: () => void;
  onSettings: () => void;
  onUserProfile: () => void;
  onSession: () => void;
  onLogOut: () => void;
  sessionData?: {
    lastLogin: string;
    currentSessionStart: string;
    sessionDuration: string;
  };
}

interface SidebarProps {
  isOpen: boolean;
  onClose: () => void;
  username: string | null;
  sessionData?: {
    lastLogin: string;
    currentSessionStart: string;
    sessionDuration: string;
  };
}

type ViewType = "main" | "devicePreference" | "settings";

// Device Preference Content Component
function DevicePreferenceContent({
  onBack,
  appParams,
  deviceParams,
}: DevicePreferenceContentProps) {
  const { t } = useTranslation();
  return (
    <div>
      {/* Header */}
      <div className="flex items-center mb-6">
        <button
          onClick={onBack}
          className="p-2 hover:bg-gray-200 rounded-lg transition-colors mr-3"
        >
          <ArrowLeft className="w-5 h-5 text-gray-600" />
        </button>
        <h1 className="text-lg font-semibold text-gray-800">
          {t("Device Preference")}
        </h1>
      </div>

      {/* App Parameters Section */}
      <div className="mb-4">
        <h2 className="text-xs font-medium text-gray-600 mb-2">
          {t("App Parameters")}
        </h2>
        <div className="bg-white rounded-lg p-3 shadow-sm">
          <div className="space-y-2">
            <div className="flex justify-between items-center">
              <span className="text-gray-700 text-sm">{t("Location")}</span>
              <span className="text-gray-900 font-medium text-sm">
                {appParams.location}
              </span>
            </div>
            <div className="flex justify-between items-center">
              <span className="text-gray-700 text-sm">
                {t("Server Reachability")}
              </span>
              <span className="text-green-600 font-medium text-sm">
                {appParams.serverReachability}
              </span>
            </div>
            <div className="flex justify-between items-center">
              <span className="text-gray-700 text-sm">
                {t("Data Connections")}
              </span>
              <span className="text-gray-900 font-medium text-sm">
                {appParams.dataConnections}
              </span>
            </div>
            <div className="flex justify-between items-center">
              <span className="text-gray-700 text-sm">{t("Latency")}</span>
              <span className="text-gray-900 font-medium text-sm">
                {appParams.latency}
              </span>
            </div>
          </div>
        </div>
      </div>

      {/* Device Parameters Section */}
      <div>
        <h2 className="text-xs font-medium text-gray-600 mb-2">
          {t("Device Parameters")}
        </h2>
        <div className="bg-white rounded-lg p-3 shadow-sm">
          <div className="space-y-2">
            <div className="flex justify-between items-start">
              <span className="text-gray-700 text-sm">{t("Platform")}</span>
              <div className="text-right">
                <div className="text-gray-900 font-medium text-sm">
                  {deviceParams.platform}
                </div>
                <div className="text-gray-500 text-xs">
                  {deviceParams.platformDetails}
                </div>
              </div>
            </div>
            <div className="flex justify-between items-center">
              <span className="text-gray-700 text-sm">{t("Manufacturer")}</span>
              <span className="text-gray-900 font-medium text-sm">
                {deviceParams.manufacturer}
              </span>
            </div>
            <div className="flex justify-between items-start">
              <span className="text-gray-700 text-sm">
                {t("Service Provider")}
              </span>
              <div className="text-right">
                <div className="text-gray-900 font-medium text-sm">
                  {deviceParams.serviceProvider}
                </div>
                <div className="text-gray-500 text-xs">
                  {deviceParams.serviceProviderDetails}
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

// Settings Content Component
function SettingsContent({ onBack }: SettingsContentProps) {
  const { t, i18n } = useTranslation();
  const [language, setLanguage] = useState<string>(i18n.language);

  const handleLanguageChange = (
    event: React.ChangeEvent<HTMLSelectElement>
  ): void => {
    const newLang = event.target.value;
    setLanguage(newLang);
    i18n.changeLanguage(newLang);
    localStorage.setItem("language", newLang);
  };

  useEffect(() => {
    const savedLang = localStorage.getItem("language");
    if (savedLang && savedLang !== i18n.language) {
      i18n.changeLanguage(savedLang);
      setLanguage(savedLang);
    }
  }, [i18n]);

  return (
    <div>
      {/* Header */}
      <div className="flex items-center mb-6">
        <button
          onClick={onBack}
          className="p-2 hover:bg-gray-200 rounded-lg transition-colors mr-3"
        >
          <ArrowLeft className="w-5 h-5 text-gray-600" />
        </button>
        <h1 className="text-lg font-semibold text-gray-800">{t("Settings")}</h1>
      </div>
      {/* Language Setting */}
      <div className="mb-4">
        <div className="bg-white rounded-lg p-3 shadow-sm">
          <div className="flex items-center justify-between">
            <div className="flex items-center space-x-3">
              <Languages className="w-4 h-4 text-gray-600" />
              <span className="text-gray-700 font-medium text-sm">
                {t("Language")}
              </span>
            </div>
            <select
              value={language}
              onChange={handleLanguageChange}
              className="bg-gray-50 border border-gray-200 rounded px-2 py-1 text-xs focus:outline-none focus:ring-1 focus:ring-blue-500"
            >
              <option value="en">{t("English(UK)")}</option>
              <option value="sv">{t("Swedish")}</option>
            </select>
          </div>
        </div>
      </div>
      {/* Additional Settings */}
      <div>
        <h2 className="text-xs font-medium text-gray-600 mb-2">
          {t("Additional Settings")}
        </h2>
        <div className="space-y-2">
          {/* <button className="w-full bg-white rounded-lg p-3 shadow-sm hover:bg-gray-50 transition-colors flex items-center justify-between">
            <span className="text-gray-700 font-medium text-sm">Notifications</span>
            <span className="text-gray-400">›</span>
          </button>
          <button className="w-full bg-white rounded-lg p-3 shadow-sm hover:bg-gray-50 transition-colors flex items-center justify-between">
            <span className="text-gray-700 font-medium text-sm">Privacy & Security</span>
            <span className="text-gray-400">›</span>
          </button>
          <button className="w-full bg-white rounded-lg p-3 shadow-sm hover:bg-gray-50 transition-colors flex items-center justify-between">
            <span className="text-gray-700 font-medium text-sm">Data Usage</span>
            <span className="text-gray-400">›</span>
          </button> */}
          <button className="w-full bg-white rounded-lg p-3 shadow-sm hover:bg-gray-50 transition-colors flex items-center justify-between">
            <span className="text-gray-700 font-medium text-sm">
              {t("About")}
            </span>
            <span className="text-gray-400">›</span>
          </button>
        </div>
      </div>
    </div>
  );
}

// Main Menu Content Component
function MainMenuContent({
  username,
  onDevicePreference,
  onSettings,
  onUserProfile,
  onSession,
  onLogOut,
  sessionData,
}: MainMenuContentProps) {
  const { t } = useTranslation();

  // Format session time for display
  const formatSessionTime = (timeString: string) => {
    if (!timeString) return t("Unknown");
    const date = new Date(timeString);
    const now = new Date();
    const diffInMinutes = Math.floor(
      (now.getTime() - date.getTime()) / (1000 * 60)
    );

    if (diffInMinutes < 1) return t("Just now");
    if (diffInMinutes < 60) return `${diffInMinutes} ${t("min ago")}`;
    if (diffInMinutes < 1440)
      return `${Math.floor(diffInMinutes / 60)} ${t("hr ago")}`;
    return `${Math.floor(diffInMinutes / 1440)} ${t("days ago")}`;
  };

  return (
    <>
      {/* Connect Section */}
      <div className="mb-8">
        <h2 className="text-gray-600 text-sm font-medium mb-4">
          TrackOpz (Connected System)
        </h2>

        {/* User Profile */}
        <button
          onClick={onUserProfile}
          className="w-full bg-green-100 rounded-xl p-4 mb-4 flex items-center justify-between hover:bg-green-200 transition-colors"
        >
          <div className="flex items-center space-x-3">
            <div className="w-8 h-8 bg-green-200 rounded-full flex items-center justify-center">
              <User className="w-4 h-4 text-green-700" />
            </div>
            <div className="text-left">
              <div className="text-gray-800 font-medium">
                {username || t("Manager")}
              </div>
            </div>
          </div>
          <span className="bg-green-600 text-white text-xs px-2 py-1 rounded-full">
            {t("Manager")}
          </span>
        </button>

        {/* Device Preference */}
        <button
          onClick={onDevicePreference}
          className="w-full flex items-center space-x-3 p-3 hover:bg-gray-200 rounded-lg transition-colors mb-2"
        >
          <Globe className="w-5 h-5 text-gray-600" />
          <span className="text-gray-700 font-medium">
            {t("Device Preference")}
          </span>
        </button>

        {/* Settings */}
        <button
          onClick={onSettings}
          className="w-full flex items-center space-x-3 p-3 hover:bg-gray-200 rounded-lg transition-colors"
        >
          <Settings className="w-5 h-5 text-gray-600" />
          <span className="text-gray-700 font-medium">{t("Settings")}</span>
        </button>
      </div>

      {/* Divider */}
      <div className="border-t border-gray-300 mb-6"></div>

      {/* Activities Section */}
      <div>
        <h2 className="text-gray-600 text-sm font-medium mb-4">
          {t("Activities")}
        </h2>

        {/* Session */}
        <button
          onClick={onSession}
          className="w-full flex items-center justify-between p-3 hover:bg-gray-200 rounded-lg transition-colors mb-2"
        >
          <div className="flex items-center space-x-3">
            <Clock className="w-5 h-5 text-gray-600" />
            <span className="text-gray-700 font-medium">{t("Session")}</span>
          </div>
          <div className="text-right">
            <div className="text-gray-500 text-sm">
              {sessionData?.lastLogin
                ? `${t("Last")}: ${formatSessionTime(sessionData.lastLogin)}`
                : t("No session data")}
            </div>
            {sessionData?.currentSessionStart && (
              <div className="text-gray-400 text-xs">
                {t("Active")}:{" "}
                {formatSessionTime(sessionData.currentSessionStart)}
              </div>
            )}
          </div>
        </button>

        {/* Log Out */}
        <button
          onClick={onLogOut}
          className="w-full flex items-center space-x-3 p-3 hover:bg-gray-200 rounded-lg transition-colors"
        >
          <LogOut className="w-5 h-5 text-gray-600" />
          <span className="text-gray-700 font-medium">{t("Log Out")}</span>
        </button>
      </div>
    </>
  );
}

export default function Sidebar({
  isOpen,
  onClose,
  username,
  sessionData,
}: SidebarProps) {
  const router = useRouter();
  const [currentView, setCurrentView] = useState<ViewType>("main");

  // Real-time App Parameters State
  const [appParams, setAppParams] = useState({
    location: "Unknown",
    serverReachability: "NO",
    dataConnections: "Unknown",
    latency: "Unknown",
  });
  const [deviceParams, setDeviceParams] = useState({
    platform: "Unknown",
    platformDetails: "",
    manufacturer: "Unknown",
    serviceProvider: "Unknown",
    serviceProviderDetails: "",
  });

  // Session data state with fallback
  const [currentSessionData, setCurrentSessionData] = useState(
    sessionData || {
      lastLogin: new Date().toISOString(),
      currentSessionStart: new Date().toISOString(),
      sessionDuration: "0 min",
    }
  );

  // Update session data when prop changes
  useEffect(() => {
    if (sessionData) {
      setCurrentSessionData(sessionData);
    }
  }, [sessionData]);

  // TODO: Replace with your real OpenCage API key
  const OPENCAGE_API_KEY = process.env.NEXT_PUBLIC_GEOCODE_API_KEY;

  useEffect(() => {
    // Platform info
    setDeviceParams((prev) => ({
      ...prev,
      platform: navigator.userAgent,
      platformDetails: navigator.platform,
    }));

    // Network info (if supported)
    const connection =
      (navigator as any).connection ||
      (navigator as any).mozConnection ||
      (navigator as any).webkitConnection;
    if (connection) {
      setAppParams((prev) => ({
        ...prev,
        dataConnections: connection.effectiveType || "Unknown",
        latency: connection.rtt ? connection.rtt + " ms" : prev.latency,
      }));
    }

    // Ping server for reachability/latency
    const start = Date.now();
    fetch("/api/ping")
      .then(() => {
        setAppParams((prev) => ({
          ...prev,
          serverReachability: "YES",
          latency: ((Date.now() - start) / 1000).toFixed(3) + " sec",
        }));
      })
      .catch(() => {
        setAppParams((prev) => ({
          ...prev,
          serverReachability: "NO",
        }));
      });

    // Geolocation with OpenCage reverse geocoding
    if (navigator.geolocation) {
      navigator.geolocation.getCurrentPosition(
        (pos) => {
          const { latitude, longitude } = pos.coords;
          // Call OpenCage API for reverse geocoding
          fetch(
            `https://api.opencagedata.com/geocode/v1/json?q=${latitude}+${longitude}&key=${OPENCAGE_API_KEY}`
          )
            .then((res) => res.json())
            .then((data) => {
              if (
                data &&
                data.results &&
                data.results[0] &&
                data.results[0].formatted
              ) {
                setAppParams((prev) => ({
                  ...prev,
                  location: data.results[0].formatted,
                }));
              } else {
                setAppParams((prev) => ({
                  ...prev,
                  location: `Lat: ${latitude.toFixed(
                    2
                  )}, Lon: ${longitude.toFixed(2)}`,
                }));
              }
            })
            .catch(() => {
              setAppParams((prev) => ({
                ...prev,
                location: `Lat: ${latitude.toFixed(
                  2
                )}, Lon: ${longitude.toFixed(2)}`,
              }));
            });
        },
        () => {
          setAppParams((prev) => ({
            ...prev,
            location: "Permission Denied",
          }));
        }
      );
    }
  }, []);

  const handleUserProfile = (): void => {
    console.log("User profile clicked");
  };

  const handleDevicePreference = (): void => {
    setCurrentView("devicePreference");
  };

  const handleSettings = (): void => {
    setCurrentView("settings");
  };

  const handleBackToMain = (): void => {
    setCurrentView("main");
  };

  const handleSession = (): void => {
    console.log("Session clicked");
  };

  const handleLogOut = async (): Promise<void> => {
    try {
      // Call the logout API to clear the JWT cookie
      const response = await fetch("/api/logout", {
        method: "POST",
        credentials: "include", // Include cookies
      });

      if (response.ok) {
        // Clear any local storage or session storage if needed
        localStorage.removeItem("language");
        sessionStorage.clear();

        // Close the sidebar
        onClose();

        // Redirect to login page
        router.push("/loginpage");
      } else {
        console.error("Logout failed");
        // Still redirect to login page even if logout API fails
        router.push("/loginpage");
      }
    } catch (error) {
      console.error("Logout error:", error);
      // Still redirect to login page even if there's an error
      router.push("/loginpage");
    }
  };

  const renderSidebarContent = () => {
    switch (currentView) {
      case "devicePreference":
        return (
          <DevicePreferenceContent
            onBack={handleBackToMain}
            appParams={appParams}
            deviceParams={deviceParams}
          />
        );
      case "settings":
        return <SettingsContent onBack={handleBackToMain} />;
      default:
        return (
          <MainMenuContent
            username={username}
            onDevicePreference={handleDevicePreference}
            onSettings={handleSettings}
            onUserProfile={handleUserProfile}
            onSession={handleSession}
            onLogOut={handleLogOut}
            sessionData={currentSessionData}
          />
        );
    }
  };

  return (
    <>
      {/* Sidebar Overlay */}
      {isOpen && (
        <div className="fixed inset-0 bg-black/50 z-40" onClick={onClose} />
      )}

      {/* Sidebar */}
      <aside
        className={`fixed top-0 left-0 h-full w-80 bg-gray-100 z-50 transform transition-transform duration-300 ${
          isOpen ? "translate-x-0" : "-translate-x-full"
        }`}
      >
        <div className="p-6 h-full overflow-y-auto">
          {/* Close Button */}
          <div className="flex justify-end mb-6">
            <button
              onClick={onClose}
              className="p-2 hover:bg-gray-200 rounded-lg transition-colors"
            >
              <X className="w-5 h-5 text-gray-600" />
            </button>
          </div>

          {/* Dynamic Content */}
          {renderSidebarContent()}
        </div>
      </aside>
    </>
  );
}
