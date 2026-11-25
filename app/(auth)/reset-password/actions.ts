"use server";

import prisma from "@/lib/prisma";
import bcrypt from "bcryptjs";
import { redirect } from "next/navigation";

export async function submitResetPassword(formData: FormData) {
    const token = formData.get("token") as string;
    const password = formData.get("password") as string;
    const konfirmasi_password = formData.get("konfirmasi_password") as string;

    if (!token || !password || !konfirmasi_password) {
        return { error: "Semua form wajib diisi!" };
    }

    if (password !== konfirmasi_password) {
        return { error: "Password dan Konfirmasi Password tidak cocok." };
    }

    try {
        // Dekode mock token (email dienkripsi base64)
        const email = Buffer.from(token, 'base64').toString('utf-8');

        const user = await prisma.user.findUnique({
            where: { email },
        });

        if (!user) {
            return { error: "Token tidak valid atau pengguna tidak ditemukan." };
        }

        const hashedPassword = await bcrypt.hash(password, 10);

        await prisma.user.update({
            where: { email },
            data: {
                password: hashedPassword,
            },
        });

    } catch (error) {
        return { error: "Terjadi kesalahan saat menyimpan password baru." };
    }

    redirect("/login?reset_success=true");
}
