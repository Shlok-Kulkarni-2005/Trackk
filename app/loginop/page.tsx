'use client';

import React, { useState } from 'react';
import { Users, User } from 'lucide-react';
import { useRouter } from 'next/navigation';

export default function LoginPage(): React.JSX.Element {
    const [phone, setPhone] = useState<string>('');
    const [otp, setOtp] = useState<string>('');
    const [otpSent, setOtpSent] = useState<boolean>(false);
    const router = useRouter();

    const handleSendOtp = (): void => {
        // Handle sending OTP logic here
        setOtpSent(true);
        console.log('OTP sent to:', phone);
    };

    const handleGetStarted = (): void => {
        router.push('/newac');
        // This should match your CSS transition duration
    };

    const handleResendOtp = (): void => {
        console.log('Resend OTP');
    };

    const handlePhoneChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
        setPhone(e.target.value);
    };

    const handleOtpChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
        setOtp(e.target.value);
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
                    <h1 className="text-4xl font-bold text-white mb-2">TrackOpz (Operator Login)</h1>
                </div>

                {/* Login Form */}
                <div className="space-y-6">
                    {/* Phone Number Field */}
                    <div className="relative">
                        <div className="absolute inset-y-0 left-0 pl-4 flex items-center pointer-events-none">
                            <User className="h-5 w-5 text-blue-300" />
                        </div>
                        <input
                            type="tel"
                            value={phone}
                            onChange={handlePhoneChange}
                            placeholder="Phone Number"
                            className="w-full pl-12 pr-4 py-4 bg-transparent border-2 border-blue-400 rounded-lg text-white placeholder-blue-300 focus:outline-none focus:border-blue-200 focus:bg-blue-600/20 transition-all duration-300"
                        />
                    </div>

                    {/* OTP Field */}
                    {otpSent && (
                        <div className="relative">
                            <input
                                type="text"
                                value={otp}
                                onChange={handleOtpChange}
                                placeholder="Enter OTP"
                                maxLength={6}
                                className="w-full pl-4 pr-4 py-4 bg-transparent border-2 border-blue-400 rounded-lg text-white placeholder-blue-300 focus:outline-none focus:border-blue-200 focus:bg-blue-600/20 transition-all duration-300"
                            />
                        </div>
                    )}

                    {/* Send OTP / Continue Button */}
                    {!otpSent ? (
                        <button
                            onClick={handleSendOtp}
                            disabled={!phone.trim()}
                            className="w-full py-4 bg-white text-blue-700 font-semibold rounded-lg hover:bg-blue-50 focus:outline-none focus:ring-4 focus:ring-blue-300/50 transition-all duration-300 transform hover:scale-[1.02] active:scale-[0.98] disabled:opacity-50 disabled:cursor-not-allowed disabled:transform-none"
                        >
                            Send OTP
                        </button>
                    ) : (
                        <button
                            onClick={handleGetStarted}
                            disabled={!otp.trim()}
                            className="w-full py-4 bg-white text-blue-700 font-semibold rounded-lg hover:bg-blue-50 focus:outline-none focus:ring-4 focus:ring-blue-300/50 transition-all duration-300 transform hover:scale-[1.02] active:scale-[0.98] disabled:opacity-50 disabled:cursor-not-allowed disabled:transform-none"
                        >
                            Continue
                        </button>
                    )}

                    {/* Resend OTP */}
                    {otpSent && (
                        <div className="text-center">
                            <button
                                type="button"
                                onClick={handleResendOtp}
                                className="text-blue-200 hover:text-blue-700 hover:bg-white transition-colors duration-300 px-2 py-1 rounded"
                            >
                                Resend OTP
                            </button>
                        </div>
                    )}
                </div>
            </div>
        </div>
    );
}