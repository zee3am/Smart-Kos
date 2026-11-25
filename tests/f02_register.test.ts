import { registerUser } from "@/app/(auth)/register/actions";
import prisma from "@/lib/prisma";
import bcrypt from "bcryptjs";
import { redirect } from "next/navigation";

jest.mock("@/lib/prisma", () => ({
  __esModule: true,
  default: {
    user: {
      findUnique: jest.fn(),
      create: jest.fn(),
    },
  },
}));

jest.mock("bcryptjs", () => ({
  hash: jest.fn(),
}));

jest.mock("next/navigation", () => ({
  redirect: jest.fn(),
}));

describe("F-02: Register Admin - Unit Testing (registerUser)", () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test("TC-REG-01 (Negatif): gagal jika salah satu input wajib kosong", async () => {
    // 1. ARRANGE
    const formData = new FormData();
    formData.append("nama", "Zidane");
    formData.append("email", ""); // kosong
    formData.append("password", "secret123");
    formData.append("no_hp", "081234567890");

    // 2. ACT
    const result = await registerUser(formData);

    // 3. ASSERT
    expect(result).toEqual({ error: "Semua form wajib diisi!" });
    expect(prisma.user.create).not.toHaveBeenCalled();
  });

  test("TC-REG-02 (Negatif): gagal jika email sudah terdaftar sebelumnya", async () => {
    // 1. ARRANGE
    const formData = new FormData();
    formData.append("nama", "Admin Baru");
    formData.append("email", "existing@smartkos.com");
    formData.append("password", "secret123");
    formData.append("no_hp", "081234567890");

    // STUB: email sudah ada
    (prisma.user.findUnique as any).mockResolvedValue({ id: "user-1", email: "existing@smartkos.com" });

    // 2. ACT
    const result = await registerUser(formData);

    // 3. ASSERT
    expect(result).toEqual({ error: "Email sudah terdaftar. Silakan gunakan email lain." });
    expect(prisma.user.create).not.toHaveBeenCalled();
  });

  test("TC-REG-03 (Positif): sukses mendaftarkan akun ADMIN dengan kode otorisasi KOSADMIN2026", async () => {
    // 1. ARRANGE
    const formData = new FormData();
    formData.append("nama", "Manager Kos");
    formData.append("email", "manager@smartkos.com");
    formData.append("password", "AdminPassword123");
    formData.append("no_hp", "081122334455");
    formData.append("role_code", "KOSADMIN2026");

    // STUB
    (prisma.user.findUnique as any).mockResolvedValue(null);
    (bcrypt.hash as any).mockResolvedValue("hashed_admin_pass");
    (prisma.user.create as any).mockResolvedValue({ id: "new-admin-id" });

    // 2. ACT
    await registerUser(formData);

    // 3. ASSERT (MOCK Verification)
    expect(bcrypt.hash).toHaveBeenCalledWith("AdminPassword123", 10);
    expect(prisma.user.create).toHaveBeenCalledWith({
      data: {
        nama: "Manager Kos",
        email: "manager@smartkos.com",
        password: "hashed_admin_pass",
        no_hp: "081122334455",
        role: "ADMIN",
      },
    });
    expect(redirect).toHaveBeenCalledWith("/login");
  });

  test("TC-REG-04 (Positif): mendaftarkan akun default role USER jika kode otorisasi tidak sesuai", async () => {
    // 1. ARRANGE
    const formData = new FormData();
    formData.append("nama", "User Reguler");
    formData.append("email", "user@smartkos.com");
    formData.append("password", "UserPassword123");
    formData.append("no_hp", "081122334499");
    formData.append("role_code", "WRONGCODE");

    (prisma.user.findUnique as any).mockResolvedValue(null);
    (bcrypt.hash as any).mockResolvedValue("hashed_user_pass");

    // 2. ACT
    await registerUser(formData);

    // 3. ASSERT
    expect(prisma.user.create).toHaveBeenCalledWith({
      data: expect.objectContaining({ role: "USER" }),
    });
    expect(redirect).toHaveBeenCalledWith("/login");
  });

  test("TC-REG-05 (Negatif): menangani error database saat eksekusi user.create gagal", async () => {
    // 1. ARRANGE
    const formData = new FormData();
    formData.append("nama", "Fail User");
    formData.append("email", "fail@smartkos.com");
    formData.append("password", "secret123");
    formData.append("no_hp", "081234567890");

    (prisma.user.findUnique as any).mockResolvedValue(null);
    (bcrypt.hash as any).mockResolvedValue("hash");
    (prisma.user.create as any).mockRejectedValue(new Error("DB Write Failure"));

    // 2. ACT
    const result = await registerUser(formData);

    // 3. ASSERT
    expect(result).toEqual({ error: "Terjadi kesalahan saat mendaftarkan akun. Coba lagi." });
  });
});
