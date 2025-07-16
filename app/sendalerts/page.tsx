"use client";
import React, { useState, ChangeEvent } from "react";
import { Menu, Send } from "lucide-react";
import Sidebar from "../../components/sidebarm";

// Type definitions
interface Alert {
  id: number;
  icon: string;
  title: string;
  description: string;
}

export default function SendAlertsPage() {
  const [message, setMessage] = useState<string>("");
  const [selectedAlert, setSelectedAlert] = useState<Alert | null>(null);
  const [sidebarOpen, setSidebarOpen] = useState<boolean>(false);
  const [notification, setNotification] = useState<string>("");

  // Sample alert data
  const alerts: Alert[] = [
    {
      id: 1,
      icon: "🔧",
      title: "Machine 1 is Under Maintenance",
      description: "Scheduled maintenance in progress",
    },
    {
      id: 2,
      icon: "⚠️",
      title: "Product A Urgent Requirement",
      description: "Critical production need",
    },
    {
      id: 3,
      icon: "🔋",
      title: "Low Power Alert",
      description: "Power levels below threshold",
    },
    {
      id: 4,
      icon: "🌡️",
      title: "Temperature Warning",
      description: "Machine overheating detected",
    },
  ];

  const handleMenuClick = (): void => {
    setSidebarOpen(true);
  };

  const handleSendAlert = async (): Promise<void> => {
    if (message.trim() || selectedAlert) {
      const alertMessage = selectedAlert ? selectedAlert.title : message;
      const alertIcon = selectedAlert ? selectedAlert.icon : "";
      try {
        const res = await fetch("/api/alerts", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ message: alertMessage, icon: alertIcon }),
        });
        if (res.ok) {
          setNotification("Alert sent successfully!");
          setMessage("");
          setSelectedAlert(null);
        } else {
          setNotification("Failed to send alert.");
        }
      } catch {
        setNotification("Failed to send alert.");
      }
      setTimeout(() => setNotification(""), 2000);
    }
  };

  const handleAlertSelect = (alert: Alert): void => {
    setSelectedAlert(alert);
    setMessage(alert.title);
  };

  const handleMessageChange = (e: ChangeEvent<HTMLTextAreaElement>): void => {
    setMessage(e.target.value);
  };

  return (
    <div className="min-h-screen bg-gray-50">
      {/* Sidebar Component */}
      <Sidebar
        isOpen={sidebarOpen}
        onClose={() => setSidebarOpen(false)}
        username="Manager"
      />

      {/* Header */}
      <header className="bg-white border-b border-gray-200 px-4 py-3">
        <div className="flex items-center justify-between">
          <button
            onClick={handleMenuClick}
            className="p-2 hover:bg-gray-100 rounded-lg"
          >
            <Menu className="w-6 h-6 text-blue-700" />
          </button>
          <h1 className="text-lg font-semibold text-blue-700">Send Alerts</h1>
          <div className="w-8 h-8 bg-indigo-600 rounded-full flex items-center justify-center">
            <span className="text-white font-medium text-sm">A</span>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="flex-1 px-4 py-6">
        {/* Quick Alert Options */}
        <div className="mb-6">
          <h2 className="text-sm font-medium text-gray-700 mb-4">
            Quick Alerts
          </h2>
          <div className="space-y-3">
            {alerts.map((alert: Alert) => (
              <button
                key={alert.id}
                onClick={() => handleAlertSelect(alert)}
                className={`w-full p-4 bg-white rounded-lg border text-left transition-all ${
                  selectedAlert?.id === alert.id
                    ? "border-indigo-500 bg-indigo-50"
                    : "border-gray-200 hover:border-gray-300"
                }`}
              >
                <div className="flex items-start space-x-3">
                  <span className="text-2xl">{alert.icon}</span>
                  <div className="flex-1">
                    <h3 className="font-medium text-gray-900 text-sm mb-1">
                      {alert.title}
                    </h3>
                    <p className="text-xs text-gray-500">{alert.description}</p>
                  </div>
                </div>
              </button>
            ))}
          </div>
        </div>

        {/* Custom Message */}
        <div className="mb-6">
          <label className="block text-sm font-medium text-gray-700 mb-3">
            Type Message
          </label>
          <div className="relative">
            <textarea
              value={message}
              onChange={handleMessageChange}
              placeholder="There is a urgent Requirement of Product A"
              className="w-full p-4 border border-gray-300 rounded-lg resize-none focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 text-sm"
              rows={3}
            />
          </div>
        </div>

        {/* Selected Alert Preview */}
        {selectedAlert && (
          <div className="mb-6 p-4 bg-indigo-50 border border-indigo-200 rounded-lg">
            <div className="flex items-center space-x-2 mb-2">
              <span className="text-lg">{selectedAlert.icon}</span>
              <span className="text-sm font-medium text-indigo-900">
                Selected Alert
              </span>
            </div>
            <p className="text-sm text-indigo-700">{selectedAlert.title}</p>
          </div>
        )}
      </main>

      {/* Send Button */}
      <div className="fixed bottom-0 left-0 right-0 bg-white border-t border-gray-200 p-4">
        {notification && (
          <div className="mb-2 text-center text-sm font-medium text-green-600">
            {notification}
          </div>
        )}
        <button
          onClick={handleSendAlert}
          disabled={!message.trim() && !selectedAlert}
          className={`w-full flex items-center justify-center space-x-2 py-3 px-4 rounded-lg font-medium text-sm transition-colors ${
            message.trim() || selectedAlert
              ? "bg-indigo-600 text-white hover:bg-indigo-700"
              : "bg-gray-300 text-gray-500 cursor-not-allowed"
          }`}
        >
          <Send className="w-4 h-4" />
          <span>Send Alert</span>
        </button>
      </div>
    </div>
  );
}
