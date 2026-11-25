"use client";

import { Card, CardBody } from "@/components/ui/legacy-card";
import { Button } from "@/components/ui/legacy-button";
import Link from "next/link";
import { useState, useTransition } from "react";
import { submitForgotPassword } from "./actions";

export default function ForgotPasswordForm() {
    const [error, setError] = useState<string | null>(null);
    const [successMsg, setSuccessMsg] = useState<{msg: string, link: string} | null>(null);
    const [isPending, startTransition] = useTransition();

    const handleSubmit = (e: React.FormEvent<HTMLFormElement>) => {
        e.preventDefault();
        setError(null);
        setSuccessMsg(null);
        
        const formData = new FormData(e.currentTarget);

        startTransition(async () => {
            const res = await submitForgotPassword(formData);
            if (res?.error) {
                setError(res.error);
            } else if (res?.success && res?.resetLink) {
                setSuccessMsg({ msg: res.success, link: res.resetLink });
            }
        });
    };

    return (
        <div className="min-h-screen bg-gray-50 flex items-center justify-center p-4">
            <Card className="w-full max-w-md shadow-lg border-none">
                <CardBody className="p-8">
                    <div className="text-center mb-8">
                        <h1 className="text-2xl font-bold text-gray-900">Lupa Password</h1>
                        <p className="text-gray-500 text-sm mt-2">Masukkan email Anda untuk menerima link reset password.</p>
                    </div>

                    {successMsg ? (
                        <div className="space-y-4">
                            <div className="p-4 bg-green-50 border border-green-200 rounded-lg text-sm text-center text-green-700">
                                <p className="font-semibold">{successMsg.msg}</p>
                                <p className="mt-2 text-gray-700">Karena ini simulasi, silakan klik tombol di bawah untuk melanjutkan reset:</p>
                            </div>
                            <Link href={successMsg.link} className="block w-full text-center bg-primary text-white pt-3 pb-3 font-bold rounded-lg shadow-md shadow-primary/20 hover:bg-primary/90 transition-colors">
                                Pergi ke Halaman Reset
                            </Link>
                        </div>
                    ) : (
                        <form onSubmit={handleSubmit} className="space-y-5">
                            {error && (
                                <div className="p-3 bg-red-50 text-red-600 border border-red-200 rounded-lg text-sm text-center">
                                    {error}
                                </div>
                            )}

                            <div className="space-y-1.5">
                                <label className="text-sm font-semibold text-gray-700">Email</label>
                                <input 
                                    name="email" 
                                    type="email" 
                                    required 
                                    className="w-full border border-gray-300 rounded-lg px-4 py-2.5 outline-none focus:ring-2 focus:ring-primary/40 focus:border-primary text-sm text-gray-900" 
                                    placeholder="email@example.com" 
                                />
                            </div>

                            <Button type="submit" variant="primary" className="w-full pt-3 pb-3 font-bold mt-2 shadow-md shadow-primary/20" disabled={isPending}>
                                {isPending ? "Memproses..." : "Kirim Link Reset"}
                            </Button>

                            <p className="text-center text-sm text-gray-500 mt-6 border-t border-gray-100 pt-6">
                                Ingat password Anda? <Link href="/login" className="text-primary font-semibold hover:underline">Kembali ke Login</Link>
                            </p>
                        </form>
                    )}
                </CardBody>
            </Card>
        </div>
    );
}
