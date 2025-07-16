"use client";

import React, { useState, ChangeEvent } from "react";
import { useRouter } from "next/navigation";

export default function AvatarPage(): React.ReactElement {
  const [username, setUsername] = useState<string>("");
  const [avatar, setAvatar] = useState<File | null>(null);
  const [avatarUrl, setAvatarUrl] = useState<string>("");
  const router = useRouter();

  const handleImageChange = (e: ChangeEvent<HTMLInputElement>): void => {
    const file = e.target.files?.[0] || null;
    setAvatar(file);
    if (file) {
      setAvatarUrl(URL.createObjectURL(file));
    } else {
      setAvatarUrl("");
    }
  };

  const goto = (): void => {
    router.push("/home2");
    // This should match your CSS transition duration
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-600 via-blue-700 to-blue-800 flex items-center justify-center p-4">
      <div className="w-full max-w-md bg-white/10 rounded-lg p-8 shadow-lg flex flex-col items-center">
        <h1 className="text-3xl font-bold text-white mb-6">
          Create Your Profile
        </h1>
        <div className="mb-6">
          <label className="block mb-2 text-white font-semibold">
            Avatar Image
          </label>
          <input
            type="file"
            accept="image/*"
            onChange={handleImageChange}
            className="block w-full text-white"
          />
          {avatarUrl && (
            <img
              src={avatarUrl}
              alt="Avatar Preview"
              className="mt-4 w-24 h-24 rounded-full object-cover border-2 border-blue-400"
            />
          )}
        </div>
        <div className="mb-6 w-full">
          <label className="block mb-2 text-white font-semibold">
            Username
          </label>
          <input
            type="text"
            value={username}
            onChange={(e: ChangeEvent<HTMLInputElement>) =>
              setUsername(e.target.value)
            }
            placeholder="Enter your username"
            className="w-full px-4 py-3 rounded-lg border-2 border-blue-400 bg-transparent text-white placeholder-blue-200 focus:outline-none focus:border-blue-200"
          />
        </div>
        <button
          onClick={goto}
          className="w-full py-3 bg-white text-blue-700 font-semibold rounded-lg hover:bg-blue-50 transition-all duration-300"
          disabled={!username || !avatar}
        >
          Continue
        </button>
      </div>
    </div>
  );
}
