import { tambahPenghuniManual, hapusPenghuni } from "@/app/admin/penghuni/actions";
import prisma from "@/lib/prisma";
import bcrypt from "bcryptjs";
import { revalidatePath } from "next/cache";
import { redirect } from "next/navigation";

jest.mock("@/lib/prisma", () => ({
  __esModule: true,
  default: {
    user: {
      findUnique: jest.fn(),
      create: jest.fn(),
      delete: jest.fn(),
    },
    kamar: {
      findUnique: jest.fn(),
      update: jest.fn(),
    },
    booking: {
      create: jest.fn(),
      deleteMany: jest.fn(),
    },
    tiket: {
      deleteMany: jest.fn(),
    },
    transaksi: {
      deleteMany: jest.fn(),
    },
  },
}));

jest.mock("bcryptjs", () => ({
  hash: jest.fn(),
}));

jest.mock("next/cache", () => ({
  revalidatePath: jest.fn(),
}));

jest.mock("next/navigation", () => ({
  redirect: jest.fn(),
}));

describe("F-08: Kelola Data Penghuni Kos - Unit Testing (tambahPenghuniManual & hapusPenghuni)", () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  describe("tambahPenghuniManual()", () => {
    test("TC-TNT-01 (Negatif): gagal jika form tidak lengkap", async () => {
      const formData = new FormData();
      formData.append("nama", "Penghuni Baru");
      formData.append("email", "");

      const result = await tambahPenghuniManual(formData);

      expect(result).toEqual({ error: "Semua kolom wajib diisi." });
      expect(prisma.user.create).not.toHaveBeenCalled();
    });

    test("TC-TNT-02 (Negatif): gagal jika email sudah terdaftar di sistem", async () => {
      const formData = new FormData();
      formData.append("nama", "Penghuni Baru");
      formData.append("email", "registered@smartkos.com");
      formData.append("no_hp", "081234567890");
      formData.append("kamar_id", "kmr-1");
      formData.append("durasi_sewa", "3");

      (prisma.user.findUnique as any).mockResolvedValue({ id: "user-reg" });

      const result = await tambahPenghuniManual(formData);

      expect(result).toEqual({ error: "Email ini sudah terdaftar di sistem." });
    });

    test("TC-TNT-03 (Negatif): gagal jika kamar tidak kosong atau tidak ditemukan", async () => {
      const formData = new FormData();
      formData.append("nama", "Penghuni Baru");
      formData.append("email", "new@smartkos.com");
      formData.append("no_hp", "081234567890");
      formData.append("kamar_id", "kmr-terisi");
      formData.append("durasi_sewa", "2");

      (prisma.user.findUnique as any).mockResolvedValue(null);
      (prisma.kamar.findUnique as any).mockResolvedValue({ id: "kmr-terisi", status: "TERISI" });

      const result = await tambahPenghuniManual(formData);

      expect(result).toEqual({ error: "Kamar yang dipilih tidak tersedia atau sudah terisi." });
    });

    test("TC-TNT-04 (Positif): berhasil menambahkan penghuni manual, alokasi kamar TERISI dan buat booking APPROVED", async () => {
      const formData = new FormData();
      formData.append("nama", "Budi Santoso");
      formData.append("email", "budi@smartkos.com");
      formData.append("no_hp", "081234567890");
      formData.append("kamar_id", "kmr-101");
      formData.append("durasi_sewa", "6");

      (prisma.user.findUnique as any).mockResolvedValue(null);
      (prisma.kamar.findUnique as any).mockResolvedValue({
        id: "kmr-101",
        status: "KOSONG",
        harga_per_bulan: 1000000,
      });
      (bcrypt.hash as any).mockResolvedValue("hashed_pass");
      (prisma.user.create as any).mockResolvedValue({ id: "user-budi" });
      (prisma.booking.create as any).mockResolvedValue({ id: "bk-budi" });
      (prisma.kamar.update as any).mockResolvedValue({ id: "kmr-101", status: "TERISI" });

      await tambahPenghuniManual(formData);

      expect(prisma.user.create).toHaveBeenCalledWith({
        data: expect.objectContaining({
          nama: "Budi Santoso",
          email: "budi@smartkos.com",
          role: "USER",
          status: "PENGHUNI",
          kamar_id: "kmr-101",
        }),
      });
      expect(prisma.booking.create).toHaveBeenCalledWith({
        data: expect.objectContaining({
          user_id: "user-budi",
          kamar_id: "kmr-101",
          status: "APPROVED",
          durasi_sewa: 6,
          total_harga: 6000000,
        }),
      });
      expect(prisma.kamar.update).toHaveBeenCalledWith({
        where: { id: "kmr-101" },
        data: { status: "TERISI" },
      });
      expect(redirect).toHaveBeenCalledWith("/admin/penghuni");
    });

    test("TC-TNT-05 (Negatif): menangani exception error saat tambah penghuni gagal", async () => {
      const formData = new FormData();
      formData.append("nama", "Error User");
      formData.append("email", "erroruser@smartkos.com");
      formData.append("no_hp", "081234567890");
      formData.append("kamar_id", "kmr-101");
      formData.append("durasi_sewa", "1");

      (prisma.user.findUnique as any).mockResolvedValue(null);
      (prisma.kamar.findUnique as any).mockResolvedValue({
        id: "kmr-101",
        status: "KOSONG",
        harga_per_bulan: 1000000,
      });
      (bcrypt.hash as any).mockResolvedValue("hashed_pass");
      (prisma.user.create as any).mockRejectedValue(new Error("Database Failure"));

      const result = await tambahPenghuniManual(formData);

      expect(result).toEqual({ error: "Gagal menambahkan penghuni. Silakan coba lagi." });
    });
  });

  describe("hapusPenghuni()", () => {
    test("TC-TNT-06 (Negatif): gagal hapus jika penghuni tidak ditemukan", async () => {
      (prisma.user.findUnique as any).mockResolvedValue(null);

      const result = await hapusPenghuni("non-existent-user");

      expect(result).toEqual({ error: "Penghuni tidak ditemukan." });
    });

    test("TC-TNT-07 (Positif): berhasil menghapus penghuni dan mengosongkan kamar yang ditempati", async () => {
      const mockUser = { id: "user-123", kamar_id: "kamar-456" };
      (prisma.user.findUnique as any).mockResolvedValue(mockUser);
      (prisma.kamar.update as any).mockResolvedValue({ id: "kamar-456", status: "KOSONG" });

      await hapusPenghuni("user-123");

      expect(prisma.kamar.update).toHaveBeenCalledWith({
        where: { id: "kamar-456" },
        data: { status: "KOSONG" },
      });
      expect(prisma.booking.deleteMany).toHaveBeenCalledWith({ where: { user_id: "user-123" } });
      expect(prisma.tiket.deleteMany).toHaveBeenCalledWith({ where: { user_id: "user-123" } });
      expect(prisma.transaksi.deleteMany).toHaveBeenCalledWith({ where: { user_id: "user-123" } });
      expect(prisma.user.delete).toHaveBeenCalledWith({ where: { id: "user-123" } });
      expect(revalidatePath).toHaveBeenCalledWith("/admin/penghuni");
    });
  });
});
