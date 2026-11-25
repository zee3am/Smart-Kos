import { submitForgotPassword } from "@/app/(auth)/forgot-password/actions";
import { submitResetPassword } from "@/app/(auth)/reset-password/actions";
import prisma from "@/lib/prisma";
import bcrypt from "bcryptjs";
import { redirect } from "next/navigation";

jest.mock("@/lib/prisma", () => ({
  __esModule: true,
  default: {
    user: {
      findUnique: jest.fn(),
      update: jest.fn(),
    },
  },
}));

jest.mock("bcryptjs", () => ({
  hash: jest.fn(),
}));

jest.mock("next/navigation", () => ({
  redirect: jest.fn(),
}));

describe("F-03: Forgot & Reset Password Admin - Unit Testing", () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  describe("submitForgotPassword()", () => {
    test("TC-FP-01 (Negatif): gagal jika input email kosong", async () => {
      const formData = new FormData();
      formData.append("email", "");

      const result = await submitForgotPassword(formData);

      expect(result).toEqual({ error: "Email wajib diisi!" });
      expect(prisma.user.findUnique).not.toHaveBeenCalled();
    });

    test("TC-FP-02 (Negatif): gagal jika email tidak ditemukan di database", async () => {
      const formData = new FormData();
      formData.append("email", "unknown@smartkos.com");

      (prisma.user.findUnique as any).mockResolvedValue(null);

      const result = await submitForgotPassword(formData);

      expect(result).toEqual({ error: "Akun dengan email tersebut tidak ditemukan." });
      expect(prisma.user.findUnique).toHaveBeenCalledWith({
        where: { email: "unknown@smartkos.com" },
      });
    });

    test("TC-FP-03 (Positif): berhasil menghasilkan token dan tautan reset password", async () => {
      const targetEmail = "admin@smartkos.com";
      const formData = new FormData();
      formData.append("email", targetEmail);

      (prisma.user.findUnique as any).mockResolvedValue({
        id: "admin-1",
        email: targetEmail,
        nama: "Admin Pengelola",
      });

      const result = await submitForgotPassword(formData);

      const expectedToken = Buffer.from(targetEmail).toString("base64");
      expect(result).toEqual({
        success: "Permintaan berhasil!",
        resetLink: `/reset-password?token=${expectedToken}`,
      });
    });

    test("TC-FP-04 (Negatif): menangani sistem error saat query database melempar exception", async () => {
      const formData = new FormData();
      formData.append("email", "error@smartkos.com");

      (prisma.user.findUnique as any).mockRejectedValue(new Error("Connection error"));

      const result = await submitForgotPassword(formData);

      expect(result).toEqual({ error: "Terjadi kesalahan sistem. Coba lagi nanti." });
    });
  });

  describe("submitResetPassword()", () => {
    test("TC-RP-01 (Negatif): gagal jika parameter token atau password kosong", async () => {
      const formData = new FormData();
      formData.append("token", "");
      formData.append("password", "newPass123");
      formData.append("konfirmasi_password", "newPass123");

      const result = await submitResetPassword(formData);

      expect(result).toEqual({ error: "Semua form wajib diisi!" });
    });

    test("TC-RP-02 (Negatif): gagal jika password dan konfirmasi password tidak cocok", async () => {
      const formData = new FormData();
      formData.append("token", "dummyToken");
      formData.append("password", "PassOne123");
      formData.append("konfirmasi_password", "PassTwo123");

      const result = await submitResetPassword(formData);

      expect(result).toEqual({ error: "Password dan Konfirmasi Password tidak cocok." });
    });

    test("TC-RP-03 (Negatif): gagal jika user dari token tidak ditemukan", async () => {
      const validToken = Buffer.from("notfound@smartkos.com").toString("base64");
      const formData = new FormData();
      formData.append("token", validToken);
      formData.append("password", "NewSecret123");
      formData.append("konfirmasi_password", "NewSecret123");

      (prisma.user.findUnique as any).mockResolvedValue(null);

      const result = await submitResetPassword(formData);

      expect(result).toEqual({ error: "Token tidak valid atau pengguna tidak ditemukan." });
    });

    test("TC-RP-04 (Positif): sukses memperbarui password baru dan dialihkan ke login", async () => {
      const targetEmail = "admin@smartkos.com";
      const validToken = Buffer.from(targetEmail).toString("base64");
      const formData = new FormData();
      formData.append("token", validToken);
      formData.append("password", "NewAdminSecret123");
      formData.append("konfirmasi_password", "NewAdminSecret123");

      (prisma.user.findUnique as any).mockResolvedValue({ id: "admin-1", email: targetEmail });
      (bcrypt.hash as any).mockResolvedValue("new_hashed_password");
      (prisma.user.update as any).mockResolvedValue({ id: "admin-1" });

      await submitResetPassword(formData);

      expect(bcrypt.hash).toHaveBeenCalledWith("NewAdminSecret123", 10);
      expect(prisma.user.update).toHaveBeenCalledWith({
        where: { email: targetEmail },
        data: { password: "new_hashed_password" },
      });
      expect(redirect).toHaveBeenCalledWith("/login?reset_success=true");
    });

    test("TC-RP-05 (Negatif): menangani exception error saat update password gagal", async () => {
      const targetEmail = "admin@smartkos.com";
      const validToken = Buffer.from(targetEmail).toString("base64");
      const formData = new FormData();
      formData.append("token", validToken);
      formData.append("password", "NewAdminSecret123");
      formData.append("konfirmasi_password", "NewAdminSecret123");

      (prisma.user.findUnique as any).mockResolvedValue({ id: "admin-1", email: targetEmail });
      (prisma.user.update as any).mockRejectedValue(new Error("Database write error"));

      const result = await submitResetPassword(formData);

      expect(result).toEqual({ error: "Terjadi kesalahan saat menyimpan password baru." });
    });
  });
});
