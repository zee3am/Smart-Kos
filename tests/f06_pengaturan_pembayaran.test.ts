import { updateKostProfile } from "@/app/admin/pengaturan/actions";
import prisma from "@/lib/prisma";
import { revalidatePath } from "next/cache";
import { redirect } from "next/navigation";
import { uploadFile, uploadMultipleFiles } from "@/lib/supabase-storage";

jest.mock("@/lib/prisma", () => ({
  __esModule: true,
  default: {
    kostProfile: {
      update: jest.fn(),
      create: jest.fn(),
    },
  },
}));

jest.mock("next/cache", () => ({
  revalidatePath: jest.fn(),
}));

jest.mock("next/navigation", () => ({
  redirect: jest.fn(),
}));

jest.mock("@/lib/supabase-storage", () => ({
  uploadFile: jest.fn(),
  uploadMultipleFiles: jest.fn(),
}));

describe("F-06: Verifikasi Tagihan & Pengaturan Pembayaran - Unit Testing (updateKostProfile)", () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test("TC-BILL-01 (Negatif): melempar error jika kolom wajib informasi kost kosong", async () => {
    const formData = new FormData();
    formData.append("nama_kost", "");

    await expect(updateKostProfile(formData)).rejects.toThrow("Semua kolom wajib diisi!");
  });

  test("TC-BILL-02 (Positif): berhasil memperbarui profil kost, rekening bank, dan mengunggah semua file gambar", async () => {
    const formData = new FormData();
    formData.append("id", "profile-id-1");
    formData.append("nama_kost", "Smart-Kos Modern");
    formData.append("alamat", "Jl. Raya Tlogomas No. 246");
    formData.append("nomor_kontak", "08123456789");
    formData.append("deskripsi", "Kos modern putra/putri nyaman");
    formData.append("hero_title", "Hunian Kos Terbaik Mahasiswa");
    formData.append("link_gmaps", "https://maps.google.com");
    formData.append("link_gmaps_button", "Buka Maps");
    formData.append("nama_bank", "BCA");
    formData.append("nomor_rekening", "1234567890");
    formData.append("nama_pemilik_rekening", "Pengelola SmartKos");

    // Mock Files
    const dummyLogo = new File(["logo"], "logo.png", { type: "image/png" });
    const dummyQris = new File(["qris"], "qris.png", { type: "image/png" });
    const dummyHero = new File(["hero"], "hero.png", { type: "image/png" });
    const dummySlider = new File(["slider"], "slider.png", { type: "image/png" });

    formData.append("logo_url", dummyLogo);
    formData.append("foto_qris", dummyQris);
    formData.append("foto_hero", dummyHero);
    formData.append("hero_images", dummySlider);

    (uploadFile as any).mockResolvedValue("https://storage.supabase.com/file.png");
    (uploadMultipleFiles as any).mockResolvedValue(["https://storage.supabase.com/slider.png"]);
    (prisma.kostProfile.update as any).mockResolvedValue({ id: "profile-id-1" });

    await updateKostProfile(formData);

    expect(prisma.kostProfile.update).toHaveBeenCalledWith({
      where: { id: "profile-id-1" },
      data: expect.objectContaining({
        nama_kost: "Smart-Kos Modern",
        nama_bank: "BCA",
        nomor_rekening: "1234567890",
        logo_url: "https://storage.supabase.com/file.png",
        foto_qris: "https://storage.supabase.com/file.png",
        foto_hero: "https://storage.supabase.com/file.png",
        hero_images: ["https://storage.supabase.com/slider.png"],
      }),
    });
    expect(revalidatePath).toHaveBeenCalledWith("/", "layout");
    expect(redirect).toHaveBeenCalledWith("/admin/pengaturan?success=1");
  });

  test("TC-BILL-03 (Positif): berhasil membuat profil kost & rekening baru jika ID belum ada", async () => {
    const formData = new FormData();
    formData.append("nama_kost", "Smart-Kos Baru");
    formData.append("alamat", "Jl. Bendungan Sutami");
    formData.append("nomor_kontak", "08111222333");
    formData.append("deskripsi", "Kos baru fasilitas lengkap");
    formData.append("hero_title", "Selamat Datang");
    formData.append("link_gmaps", "https://maps.google.com");
    formData.append("link_gmaps_button", "Peta Lokasi");
    formData.append("nama_bank", "Mandiri");
    formData.append("nomor_rekening", "9876543210");
    formData.append("nama_pemilik_rekening", "Ibu Kos");

    (prisma.kostProfile.create as any).mockResolvedValue({ id: "new-profile" });

    await updateKostProfile(formData);

    expect(prisma.kostProfile.create).toHaveBeenCalledWith({
      data: expect.objectContaining({
        nama_kost: "Smart-Kos Baru",
        nama_bank: "Mandiri",
        nomor_rekening: "9876543210",
      }),
    });
    expect(redirect).toHaveBeenCalledWith("/admin/pengaturan?success=1");
  });
});
