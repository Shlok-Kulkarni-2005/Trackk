'use client';
import React, { useState } from "react";
import { Mail, Eye, EyeOff, MailCheck } from "lucide-react";

// Type for the current page state
type PageState = "forgot" | "email-sent";

export default function ForgotPasswordFlow(): React.JSX.Element {
  const [emailVisible, setEmailVisible] = useState<boolean>(false);
  const [email, setEmail] = useState<string>("");
  const [error, setError] = useState<string>("");
  const [currentPage, setCurrentPage] = useState<PageState>("forgot");

  const toggleEmailVisibility = (): void => {
    setEmailVisible((prev: boolean) => !prev);
  };

  const validateEmail = (email: string): boolean => {
    // Basic email format regex
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
  };

  const handleResetClick = (): void => {
    if (!email.trim()) {
      setError("Email is required.");
      return;
    }
    if (!validateEmail(email)) {
      setError("Please enter a valid email address.");
      return;
    }
    // All good – clear error and proceed
    setError("");
    // Navigate to email sent page
    setCurrentPage("email-sent");
  };

  const handleBackToLogin = (): void => {
    // Reset form and go back to forgot password page
    setCurrentPage("forgot");
    setEmail("");
    setError("");
    setEmailVisible(false);
  };

  const handleKeyPress = (e: React.KeyboardEvent<HTMLInputElement>): void => {
    if (e.key === 'Enter') {
      handleResetClick();
    }
  };

  const handleEmailChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    setEmail(e.target.value);
  };

  if (currentPage === "email-sent") {
    return (
      <div className="min-h-screen bg-blue-700 flex flex-col justify-center items-center p-4">
        <h1 className="text-white text-2xl md:text-3xl mb-4 font-bold text-center">
          Check Your Email
        </h1>
        <div className="bg-white p-6 md:p-8 rounded-lg w-full max-w-sm text-center shadow-lg">
          <MailCheck size={32} className="text-blue-500 mx-auto mb-4" />
          <h2 className="text-blue-500 text-xl font-semibold mb-1">
            Reset Link Sent
          </h2>
          <p className="text-gray-600 text-sm mb-6">
            If your email is registered, we've sent a link to reset your password.
            Please check your inbox.
          </p>
          <button
            onClick={handleBackToLogin}
            className="bg-blue-500 text-white border-none w-full py-2.5 rounded-md font-semibold cursor-pointer hover:bg-blue-600 transition-colors"
            type="button"
          >
            Back to Login
          </button>
        </div>
        <footer className="mt-6 text-white text-xs text-center">
          © 2025 Reset Password Form. All rights reserved | TrackOpz
        </footer>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-blue-700 flex flex-col justify-center items-center p-4">
      <h1 className="text-white text-2xl md:text-3xl mb-4 font-bold text-center">
        Reset Your Password
      </h1>
      <div className="bg-white p-6 md:p-8 rounded-lg w-full max-w-sm text-center shadow-lg">
        <h2 className="text-blue-500 text-xl font-semibold mb-1">
          Forgot Password
        </h2>
        <p className="text-gray-600 text-sm mb-6">
          Please enter your registered email address.
        </p>
        <div className="flex items-center border border-gray-300 p-2 rounded-md mb-4">
          <Mail size={16} className="mr-2 text-gray-700" />
          <input
            type={emailVisible ? "text" : "password"}
            placeholder="Your e-mail address"
            className="border-none outline-none flex-grow text-sm text-gray-700"
            value={email}
            onChange={handleEmailChange}
            onKeyDown={handleKeyPress}
            aria-label="Email address"
          />
          <button
            type="button"
            onClick={toggleEmailVisibility}
            className="ml-auto text-gray-700 cursor-pointer hover:text-gray-900 p-1"
            aria-label={emailVisible ? "Hide email" : "Show email"}
          >
            {emailVisible ? (
              <EyeOff size={16} />
            ) : (
              <Eye size={16} />
            )}
          </button>
        </div>
        {error && (
          <p className="text-red-500 text-xs mb-4 text-left" role="alert">
            {error}
          </p>
        )}
        <button
          onClick={handleResetClick}
          className="bg-blue-500 text-white border-none w-full py-2.5 rounded-md font-semibold cursor-pointer hover:bg-blue-600 transition-colors"
          type="button"
        >
          Reset my Password
        </button>
      </div>
      {/* <footer className="mt-6 text-white text-xs text-center">
        © 2025 Reset Password Form. All rights reserved | TrackOpz
      </footer> */}
    </div>
  );
}