import { approveBooking, rejectBooking } from "@/app/admin/verifikasi/actions";
import prisma from "@/lib/prisma";
import { revalidatePath } from "next/cache";

// ============================================================================
// MOCK & STUB CONFIGURATION
// Mengisolasi dependensi eksternal (Prisma ORM & Next.js Cache)
// ============================================================================
jest.mock("@/lib/prisma", () => ({
  __esModule: true,
  default: {
    booking: {
      findUnique: jest.fn(),
      findFirst: jest.fn(),
      update: jest.fn(),
    },
    kamar: {
      update: jest.fn(),
    },
    user: {
      update: jest.fn(),
    },
  },
}));

jest.mock("next/cache", () => ({
  revalidatePath: jest.fn(),
}));

describe("Unit Testing: Feature Verifikasi Booking Admin (app/admin/verifikasi/actions)", () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  // ==========================================================================
  // SKENARIO 1 (NEGATIF): Booking ID Tidak Ditemukan
  // Menguji penanganan error ketika data booking tidak ada di database
  // ==========================================================================
  test("TC-01: harus mengembalikan error jika bookingId tidak ditemukan di database", async () => {
    // --------------------------------------------------
    // 1. ARRANGE
    // --------------------------------------------------
    const invalidBookingId = "booking-invalid-999";

    // TEST DOUBLE 1 (STUB):
    // Mensimulasikan prisma.booking.findUnique mengembalikan null (canned response)
    (prisma.booking.findUnique as jest.Mock).mockResolvedValue(null);

    // --------------------------------------------------
    // 2. ACT
    // --------------------------------------------------
    const result = await approveBooking(invalidBookingId);

    // --------------------------------------------------
    // 3. ASSERT
    // --------------------------------------------------
    // Validasi nilai return error
    expect(result).toEqual({ error: "Pesanan tidak ditemukan" });

    // TEST DOUBLE 2 (MOCK):
    // Verifikasi interaksi: pastikan findUnique dipanggil dengan ID yang sesuai
    expect(prisma.booking.findUnique).toHaveBeenCalledWith({
      where: { id: invalidBookingId },
    });

    // Pastikan tidak ada aksi update database yang dieksekusi
    expect(prisma.booking.update).not.toHaveBeenCalled();
    expect(prisma.kamar.update).not.toHaveBeenCalled();
    expect(prisma.user.update).not.toHaveBeenCalled();
    expect(revalidatePath).not.toHaveBeenCalled();
  });

  // ==========================================================================
  // SKENARIO 2 (POSITIF): Sewa Baru (Tanpa Perpanjangan Booking Aktif)
  // Menguji persetujuan sewa awal penghuni baru (dihitung mulai hari ini)
  // ==========================================================================
  test("TC-02: harus menyetujui booking baru dan menghitung tanggal berakhir dari hari ini jika tidak ada sewa aktif sebelumnya", async () => {
    // --------------------------------------------------
    // 1. ARRANGE
    // --------------------------------------------------
    const bookingId = "bk-new-001";
    const dummyBooking = {
      id: bookingId,
      user_id: "user-alpha",
      kamar_id: "kamar-101",
      durasi_sewa: 1, // 1 bulan
      status: "PENDING",
    };

    // TEST DOUBLE 1 (STUB):
    // Mengembalikan data booking valid dan lastActiveBooking null (tidak ada perpanjangan)
    (prisma.booking.findUnique as jest.Mock).mockResolvedValue(dummyBooking);
    (prisma.booking.findFirst as jest.Mock).mockResolvedValue(null);
    (prisma.booking.update as jest.Mock).mockResolvedValue({ id: bookingId, status: "APPROVED" });
    (prisma.kamar.update as jest.Mock).mockResolvedValue({ id: dummyBooking.kamar_id, status: "TERISI" });
    (prisma.user.update as jest.Mock).mockResolvedValue({ id: dummyBooking.user_id, status: "PENGHUNI" });

    const beforeCallTime = new Date();

    // --------------------------------------------------
    // 2. ACT
    // --------------------------------------------------
    const result = await approveBooking(bookingId);

    // --------------------------------------------------
    // 3. ASSERT
    // --------------------------------------------------
    // Fungsi berhasil menyelesaikan seluruh proses (void return)
    expect(result).toBeUndefined();

    // TEST DOUBLE 2 (MOCK):
    // Verifikasi pemanggilan prisma.booking.update dengan data APPROVED dan tanggal_berakhir
    expect(prisma.booking.update).toHaveBeenCalledTimes(1);
    const updateCallArg = (prisma.booking.update as jest.Mock).mock.calls[0][0];
    expect(updateCallArg.where).toEqual({ id: bookingId });
    expect(updateCallArg.data.status).toBe("APPROVED");
    
    // Periksa bahwa tanggal berakhir dihitung ~1 bulan dari hari ini
    const calculatedEndDate = new Date(updateCallArg.data.tanggal_berakhir);
    expect(calculatedEndDate.getTime()).toBeGreaterThan(beforeCallTime.getTime());

    // Verifikasi mutasi status kamar menjadi TERISI
    expect(prisma.kamar.update).toHaveBeenCalledWith({
      where: { id: dummyBooking.kamar_id },
      data: { status: "TERISI" },
    });

    // Verifikasi mutasi status user menjadi PENGHUNI dan kamar_id terikat
    expect(prisma.user.update).toHaveBeenCalledWith({
      where: { id: dummyBooking.user_id },
      data: {
        status: "PENGHUNI",
        kamar_id: dummyBooking.kamar_id,
      },
    });

    // Verifikasi revalidasi path Next.js
    expect(revalidatePath).toHaveBeenCalledWith("/admin/verifikasi");
    expect(revalidatePath).toHaveBeenCalledWith("/admin/kamar");
  });

  // ==========================================================================
  // SKENARIO 3 (POSITIF): Perpanjangan Sewa (Terdapat Sewa Aktif Sebelumnya)
  // Menguji perpanjangan sewa: tanggal sewa baru diakumulasikan dari tanggal sewa lama
  // ==========================================================================
  test("TC-03: harus menyetujui perpanjangan sewa dan mengakumulasikan tanggal berakhir dari masa sewa aktif sebelumnya", async () => {
    // --------------------------------------------------
    // 1. ARRANGE
    // --------------------------------------------------
    const bookingId = "bk-ext-002";
    const dummyBooking = {
      id: bookingId,
      user_id: "user-beta",
      kamar_id: "kamar-102",
      durasi_sewa: 3, // Perpanjang 3 bulan
      status: "PENDING",
    };

    // Sewa aktif sebelumnya masih berlaku hingga 15 hari ke depan
    const activeEndDate = new Date();
    activeEndDate.setDate(activeEndDate.getDate() + 15);

    const dummyLastActiveBooking = {
      id: "bk-prev-active",
      user_id: "user-beta",
      status: "APPROVED",
      tanggal_berakhir: activeEndDate,
    };

    // TEST DOUBLE 1 (STUB):
    (prisma.booking.findUnique as jest.Mock).mockResolvedValue(dummyBooking);
    (prisma.booking.findFirst as jest.Mock).mockResolvedValue(dummyLastActiveBooking);
    (prisma.booking.update as jest.Mock).mockResolvedValue({ id: bookingId, status: "APPROVED" });
    (prisma.kamar.update as jest.Mock).mockResolvedValue({ id: dummyBooking.kamar_id, status: "TERISI" });
    (prisma.user.update as jest.Mock).mockResolvedValue({ id: dummyBooking.user_id, status: "PENGHUNI" });

    // --------------------------------------------------
    // 2. ACT
    // --------------------------------------------------
    const result = await approveBooking(bookingId);

    // --------------------------------------------------
    // 3. ASSERT
    // --------------------------------------------------
    expect(result).toBeUndefined();

    // TEST DOUBLE 2 (MOCK):
    // Verifikasi bahwa tanggal_berakhir dihitung dari activeEndDate + 3 bulan
    expect(prisma.booking.update).toHaveBeenCalledTimes(1);
    const updateCallArg = (prisma.booking.update as jest.Mock).mock.calls[0][0];
    const newCalculatedEndDate = new Date(updateCallArg.data.tanggal_berakhir);

    const expectedEndDate = new Date(activeEndDate);
    expectedEndDate.setMonth(expectedEndDate.getMonth() + 3);

    expect(newCalculatedEndDate.getFullYear()).toBe(expectedEndDate.getFullYear());
    expect(newCalculatedEndDate.getMonth()).toBe(expectedEndDate.getMonth());
    expect(newCalculatedEndDate.getDate()).toBe(expectedEndDate.getDate());

    // Verifikasi mutasi kamar dan user
    expect(prisma.kamar.update).toHaveBeenCalledWith({
      where: { id: dummyBooking.kamar_id },
      data: { status: "TERISI" },
    });
    expect(prisma.user.update).toHaveBeenCalledWith({
      where: { id: dummyBooking.user_id },
      data: {
        status: "PENGHUNI",
        kamar_id: dummyBooking.kamar_id,
      },
    });

    expect(revalidatePath).toHaveBeenCalledWith("/admin/verifikasi");
    expect(revalidatePath).toHaveBeenCalledWith("/admin/kamar");
  });

  // ==========================================================================
  // SKENARIO 4 & 5: rejectBooking()
  // ==========================================================================
  describe("rejectBooking()", () => {
    test("TC-04: harus mengembalikan error jika bookingId tidak ditemukan saat reject", async () => {
      (prisma.booking.findUnique as jest.Mock).mockResolvedValue(null);

      const result = await rejectBooking("bk-not-found");

      expect(result).toEqual({ error: "Pesanan tidak ditemukan" });
      expect(prisma.booking.update).not.toHaveBeenCalled();
    });

    test("TC-05: harus memperbarui status booking menjadi REJECTED ketika penolakan berhasil", async () => {
      (prisma.booking.findUnique as jest.Mock).mockResolvedValue({ id: "bk-valid-reject" });
      (prisma.booking.update as jest.Mock).mockResolvedValue({ id: "bk-valid-reject", status: "REJECTED" });

      const result = await rejectBooking("bk-valid-reject");

      expect(result).toBeUndefined();
      expect(prisma.booking.update).toHaveBeenCalledWith({
        where: { id: "bk-valid-reject" },
        data: { status: "REJECTED" },
      });
      expect(revalidatePath).toHaveBeenCalledWith("/admin/verifikasi");
    });
  });
});
