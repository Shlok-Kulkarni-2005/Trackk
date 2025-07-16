"use client";

import React, { useState } from "react";
import { User, Lock, Users } from "lucide-react";
import { useRouter } from "next/navigation";

export default function LoginPage(): React.JSX.Element {
  const [username, setUsername] = useState<string>("");
  const [password, setPassword] = useState<string>("");
  const [showAuthKey, setShowAuthKey] = useState<boolean>(false);
  const [authKey, setAuthKey] = useState<string>("");

  const router = useRouter();

  const handleSubmit = (): void => {
    // Handle login logic here
    setShowAuthKey(true);
  };

  const startone = (): void => {
    router.push("/home");
    // This should match your CSS transition duration
  };

  const handleForgotPassword = (): void => {
    router.push("/forgotpass");
  };

  const handleUsernameChange = (
    e: React.ChangeEvent<HTMLInputElement>
  ): void => {
    setUsername(e.target.value);
  };

  const handlePasswordChange = (
    e: React.ChangeEvent<HTMLInputElement>
  ): void => {
    setPassword(e.target.value);
  };

  const handleAuthKeyChange = (
    e: React.ChangeEvent<HTMLInputElement>
  ): void => {
    setAuthKey(e.target.value);
  };

  const handleOperatorLogin = (): void => {
    router.push("/loginop");
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-600 via-blue-700 to-blue-800 flex items-center justify-center p-4">
      <div className="w-full max-w-md">
        {/* Users Icon */}
        <div className="flex justify-center mb-12">
          <div className="relative">
            <Users className="w-19 h-19 text-white stroke-[1.5]" />
            <div className="absolute -top-2 -right-2 w-6 h-6 bg-white rounded-full flex items-center justify-center">
              <div className="w-2 h-2 bg-blue-600 rounded-full"></div>
            </div>
          </div>
        </div>

        {/* Title */}
        <div className="text-center mb-8">
          <h1 className="text-4xl font-bold text-white mb-2">
            TrackOpz (Manager Login)
          </h1>
        </div>

        <div className="space-y-6">
          {!showAuthKey ? (
            <>
              {/* Username Field */}
              <div className="relative">
                <div className="absolute inset-y-0 left-0 pl-4 flex items-center pointer-events-none">
                  <User className="h-5 w-5 text-blue-300" />
                </div>
                <input
                  type="text"
                  value={username}
                  onChange={handleUsernameChange}
                  placeholder="MANAGER ID"
                  className="w-full pl-12 pr-4 py-4 bg-transparent border-2 border-blue-400 rounded-lg text-white placeholder-blue-300 focus:outline-none focus:border-blue-200 focus:bg-blue-600/20 transition-all duration-300"
                />
              </div>

              {/* Password Field */}
              <div className="relative">
                <div className="absolute inset-y-0 left-0 pl-4 flex items-center pointer-events-none">
                  <Lock className="h-5 w-5 text-blue-300" />
                </div>
                <input
                  type="password"
                  value={password}
                  onChange={handlePasswordChange}
                  placeholder="PASSWORD"
                  className="w-full pl-12 pr-4 py-4 bg-transparent border-2 border-blue-400 rounded-lg text-white placeholder-blue-300 focus:outline-none focus:border-blue-200 focus:bg-blue-600/20 transition-all duration-300"
                />
              </div>

              {/* Login Button */}
              <button
                onClick={handleSubmit}
                disabled={!username.trim() || !password.trim()}
                className="w-full py-4 bg-white text-blue-700 font-semibold rounded-lg hover:bg-blue-50 focus:outline-none focus:ring-4 focus:ring-blue-300/50 transition-all duration-300 transform hover:scale-[1.02] active:scale-[0.98] disabled:opacity-50 disabled:cursor-not-allowed disabled:transform-none"
              >
                LOGIN
              </button>

              {/* Forgot Password Link */}
              <div className="text-right">
                <button
                  type="button"
                  onClick={handleForgotPassword}
                  className="text-blue-200 hover:text-white transition-colors duration-300"
                >
                  Forgot password?
                </button>
              </div>
            </>
          ) : (
            <>
              {/* Auth Key Field */}
              <div className="relative">
                <input
                  type="text"
                  value={authKey}
                  onChange={handleAuthKeyChange}
                  placeholder="ENTER AUTHKEY"
                  className="w-full pl-4 pr-4 py-4 bg-transparent border-2 border-blue-400 rounded-lg text-white placeholder-blue-300 focus:outline-none focus:border-blue-200 focus:bg-blue-600/20 transition-all duration-300"
                />
              </div>

              {/* Continue Button */}
              <button
                onClick={startone}
                disabled={!authKey.trim()}
                className="w-full py-4 bg-white text-blue-700 font-semibold rounded-lg hover:bg-blue-50 focus:outline-none focus:ring-4 focus:ring-blue-300/50 transition-all duration-300 transform hover:scale-[1.02] active:scale-[0.98] disabled:opacity-50 disabled:cursor-not-allowed disabled:transform-none"
              >
                CONTINUE
              </button>
            </>
          )}

          {/* Operator Login */}
          <div className="text-center">
            <button
              type="button"
              onClick={handleOperatorLogin}
              className="text-blue-200 hover:text-blue-700 hover:bg-white transition-colors duration-300 px-2 py-1 rounded"
            >
              Operator Login
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
