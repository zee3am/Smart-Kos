"use server";

import prisma from "@/lib/prisma";

export async function submitForgotPassword(formData: FormData) {
    const email = formData.get("email") as string;

    if (!email) {
        return { error: "Email wajib diisi!" };
    }

    try {
        const user = await prisma.user.findUnique({
            where: { email },
        });

        if (!user) {
            // Untuk keamanan, sebaiknya pesannya ambigu, tapi karena ini untuk keperluan modul 1:
            return { error: "Akun dengan email tersebut tidak ditemukan." };
        }

        // Simulasi: Mengembalikan link reset yang berisi email pengguna sebagai mock-token
        const mockToken = Buffer.from(email).toString('base64');
        const resetLink = `/reset-password?token=${mockToken}`;

        return { 
            success: "Permintaan berhasil!", 
            resetLink 
        };

    } catch (error) {
        return { error: "Terjadi kesalahan sistem. Coba lagi nanti." };
    }
}
