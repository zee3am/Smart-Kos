"use client";

import { Card, CardBody } from "@/components/ui/legacy-card";
import { Button } from "@/components/ui/legacy-button";
import { useState, useTransition, Suspense } from "react";
import { useSearchParams } from "next/navigation";
import { submitResetPassword } from "./actions";

function ResetPasswordFormContent() {
    const searchParams = useSearchParams();
    const token = searchParams.get("token");
    const [error, setError] = useState<string | null>(null);
    const [isPending, startTransition] = useTransition();
    const [showPassword, setShowPassword] = useState(false);
    const [showConfirmPassword, setShowConfirmPassword] = useState(false);

    if (!token) {
        return (
            <div className="text-center">
                <p className="text-red-600 font-semibold mb-4">Token reset password tidak valid atau tidak ditemukan.</p>
                <a href="/forgot-password" className="text-primary hover:underline">Kembali ke Lupa Password</a>
            </div>
        );
    }

    const handleSubmit = (e: React.FormEvent<HTMLFormElement>) => {
        e.preventDefault();
        setError(null);
        
        const formData = new FormData(e.currentTarget);
        // Tambahkan token dari URL ke formData
        formData.append("token", token);

        startTransition(async () => {
            const res = await submitResetPassword(formData);
            if (res?.error) {
                setError(res.error);
            }
        });
    };

    return (
        <form onSubmit={handleSubmit} className="space-y-5">
            {error && (
                <div className="p-3 bg-red-50 text-red-600 border border-red-200 rounded-lg text-sm text-center">
                    {error}
                </div>
            )}

            <div className="space-y-1.5">
                <label className="text-sm font-semibold text-gray-700">Password Baru</label>
                <div className="relative">
                    <input
                        name="password"
                        type={showPassword ? "text" : "password"}
                        required
                        className="w-full border border-gray-300 rounded-lg px-4 py-2.5 outline-none focus:ring-2 focus:ring-primary/40 focus:border-primary text-sm text-gray-900 pr-10"
                        placeholder="••••••••"
                    />
                    <button
                        type="button"
                        className="absolute inset-y-0 right-0 pr-3 flex items-center text-gray-500 hover:text-gray-700"
                        onClick={() => setShowPassword(!showPassword)}
                    >
                        {showPassword ? "Hide" : "Show"}
                    </button>
                </div>
            </div>

            <div className="space-y-1.5">
                <label className="text-sm font-semibold text-gray-700">Konfirmasi Password Baru</label>
                <div className="relative">
                    <input
                        name="konfirmasi_password"
                        type={showConfirmPassword ? "text" : "password"}
                        required
                        className="w-full border border-gray-300 rounded-lg px-4 py-2.5 outline-none focus:ring-2 focus:ring-primary/40 focus:border-primary text-sm text-gray-900 pr-10"
                        placeholder="••••••••"
                    />
                    <button
                        type="button"
                        className="absolute inset-y-0 right-0 pr-3 flex items-center text-gray-500 hover:text-gray-700"
                        onClick={() => setShowConfirmPassword(!showConfirmPassword)}
                    >
                        {showConfirmPassword ? "Hide" : "Show"}
                    </button>
                </div>
            </div>

            <Button type="submit" variant="primary" className="w-full pt-3 pb-3 font-bold mt-2 shadow-md shadow-primary/20" disabled={isPending}>
                {isPending ? "Menyimpan..." : "Simpan Password Baru"}
            </Button>
        </form>
    );
}

export default function ResetPasswordPage() {
    return (
        <div className="min-h-screen bg-gray-50 flex items-center justify-center p-4">
            <Card className="w-full max-w-md shadow-lg border-none">
                <CardBody className="p-8">
                    <div className="text-center mb-8">
                        <h1 className="text-2xl font-bold text-gray-900">Reset Password</h1>
                        <p className="text-gray-500 text-sm mt-2">Buat kata sandi baru untuk akun Anda.</p>
                    </div>

                    <Suspense fallback={<div className="text-center text-gray-500">Memuat form...</div>}>
                        <ResetPasswordFormContent />
                    </Suspense>
                </CardBody>
            </Card>
        </div>
    );
}
