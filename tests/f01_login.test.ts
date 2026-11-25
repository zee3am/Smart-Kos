import { loginUser } from "@/app/(auth)/login/actions";
import prisma from "@/lib/prisma";
import bcrypt from "bcryptjs";
import { redirect } from "next/navigation";
import { createSession } from "@/lib/auth";

jest.mock("@/lib/prisma", () => ({
  __esModule: true,
  default: {
    user: {
      findUnique: jest.fn(),
    },
  },
}));

jest.mock("bcryptjs", () => ({
  compare: jest.fn(),
}));

jest.mock("next/navigation", () => ({
  redirect: jest.fn(),
}));

jest.mock("@/lib/auth", () => ({
  createSession: jest.fn(),
}));

describe("F-01: Login Admin - Unit Testing (loginUser)", () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test("TC-AUTH-01 (Negatif): gagal login jika form email atau password kosong", async () => {
    // 1. ARRANGE
    const formData = new FormData();
    formData.append("email", "");
    formData.append("password", "");

    // 2. ACT
    const result = await loginUser(formData);

    // 3. ASSERT
    expect(result).toEqual({ error: "Semua form wajib diisi!" });
    expect(prisma.user.findUnique).not.toHaveBeenCalled();
  });

  test("TC-AUTH-02 (Negatif): gagal login jika email tidak ditemukan di database", async () => {
    // 1. ARRANGE
    const formData = new FormData();
    formData.append("email", "notfound@smartkos.com");
    formData.append("password", "Password123");

    // STUB: user null
    (prisma.user.findUnique as any).mockResolvedValue(null);

    // 2. ACT
    const result = await loginUser(formData);

    // 3. ASSERT
    expect(result).toEqual({ error: "Email atau password salah." });
    expect(prisma.user.findUnique).toHaveBeenCalledWith({
      where: { email: "notfound@smartkos.com" },
    });
  });

  test("TC-AUTH-03 (Negatif): gagal login jika password salah (bcrypt mismatch)", async () => {
    // 1. ARRANGE
    const formData = new FormData();
    formData.append("email", "admin@smartkos.com");
    formData.append("password", "WrongPassword");

    // STUB: user ditemukan tapi bcrypt false
    (prisma.user.findUnique as any).mockResolvedValue({
      id: "admin-1",
      email: "admin@smartkos.com",
      password: "hashed_password",
      role: "ADMIN",
      status: "PENGHUNI",
    });
    (bcrypt.compare as any).mockResolvedValue(false);

    // 2. ACT
    const result = await loginUser(formData);

    // 3. ASSERT
    expect(result).toEqual({ error: "Email atau password salah." });
    expect(createSession).not.toHaveBeenCalled();
  });

  test("TC-AUTH-04 (Positif): sukses login sebagai ADMIN dan dialihkan ke /admin", async () => {
    // 1. ARRANGE
    const formData = new FormData();
    formData.append("email", "admin@smartkos.com");
    formData.append("password", "AdminPassword123");

    const mockAdminUser = {
      id: "admin-id-01",
      email: "admin@smartkos.com",
      password: "hashed_secret",
      role: "ADMIN",
      status: "PENGHUNI",
    };

    // STUB
    (prisma.user.findUnique as any).mockResolvedValue(mockAdminUser);
    (bcrypt.compare as any).mockResolvedValue(true);
    (createSession as any).mockResolvedValue(undefined);

    // 2. ACT
    await loginUser(formData);

    // 3. ASSERT (MOCK Verification)
    expect(createSession).toHaveBeenCalledWith(mockAdminUser.id, "ADMIN", "PENGHUNI");
    expect(redirect).toHaveBeenCalledWith("/admin");
  });

  test("TC-AUTH-05 (Positif): sukses login sebagai USER biasa dan dialihkan ke /dashboard", async () => {
    // 1. ARRANGE
    const formData = new FormData();
    formData.append("email", "user@smartkos.com");
    formData.append("password", "UserPassword123");

    const mockUser = {
      id: "user-id-02",
      email: "user@smartkos.com",
      password: "hashed_secret",
      role: "USER",
      status: "PENGHUNI",
    };

    (prisma.user.findUnique as any).mockResolvedValue(mockUser);
    (bcrypt.compare as any).mockResolvedValue(true);

    // 2. ACT
    await loginUser(formData);

    // 3. ASSERT
    expect(createSession).toHaveBeenCalledWith(mockUser.id, "USER", "PENGHUNI");
    expect(redirect).toHaveBeenCalledWith("/dashboard");
  });

  test("TC-AUTH-06 (Negatif): menangani sistem error saat query database melempar exception", async () => {
    // 1. ARRANGE
    const formData = new FormData();
    formData.append("email", "error@smartkos.com");
    formData.append("password", "Password123");

    // STUB exception
    (prisma.user.findUnique as any).mockRejectedValue(new Error("Database connection timeout"));

    // 2. ACT
    const result = await loginUser(formData);

    // 3. ASSERT
    expect(result).toEqual({ error: "Terjadi kesalahan sistem. Coba lagi nanti." });
  });
});
