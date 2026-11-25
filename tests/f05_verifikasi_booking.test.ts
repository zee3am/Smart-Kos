import { approveBooking, rejectBooking } from "@/app/admin/verifikasi/actions";
import prisma from "@/lib/prisma";
import { revalidatePath } from "next/cache";

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

describe("F-05: Verifikasi Booking Kamar - Unit Testing (approveBooking & rejectBooking)", () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test("TC-BOOK-01 (Negatif): gagal jika bookingId tidak ditemukan di database", async () => {
    // 1. ARRANGE
    const invalidBookingId = "booking-invalid-999";
    (prisma.booking.findUnique as any).mockResolvedValue(null);

    // 2. ACT
    const result = await approveBooking(invalidBookingId);

    // 3. ASSERT
    expect(result).toEqual({ error: "Pesanan tidak ditemukan" });
    expect(prisma.booking.findUnique).toHaveBeenCalledWith({ where: { id: invalidBookingId } });
    expect(prisma.booking.update).not.toHaveBeenCalled();
  });

  test("TC-BOOK-02 (Positif): menyetujui booking baru dan menghitung tanggal berakhir dari hari ini", async () => {
    // 1. ARRANGE
    const bookingId = "bk-new-001";
    const dummyBooking = {
      id: bookingId,
      user_id: "user-alpha",
      kamar_id: "kamar-101",
      durasi_sewa: 1,
      status: "PENDING",
    };

    (prisma.booking.findUnique as any).mockResolvedValue(dummyBooking);
    (prisma.booking.findFirst as any).mockResolvedValue(null);
    (prisma.booking.update as any).mockResolvedValue({ id: bookingId, status: "APPROVED" });
    (prisma.kamar.update as any).mockResolvedValue({ id: dummyBooking.kamar_id, status: "TERISI" });
    (prisma.user.update as any).mockResolvedValue({ id: dummyBooking.user_id, status: "PENGHUNI" });

    // 2. ACT
    const result = await approveBooking(bookingId);

    // 3. ASSERT
    expect(result).toBeUndefined();
    expect(prisma.booking.update).toHaveBeenCalledTimes(1);
    expect(prisma.kamar.update).toHaveBeenCalledWith({
      where: { id: dummyBooking.kamar_id },
      data: { status: "TERISI" },
    });
    expect(prisma.user.update).toHaveBeenCalledWith({
      where: { id: dummyBooking.user_id },
      data: { status: "PENGHUNI", kamar_id: dummyBooking.kamar_id },
    });
    expect(revalidatePath).toHaveBeenCalledWith("/admin/verifikasi");
    expect(revalidatePath).toHaveBeenCalledWith("/admin/kamar");
  });

  test("TC-BOOK-03 (Positif): menyetujui perpanjangan sewa booking aktif dari tanggal berakhir sewa sebelumnya", async () => {
    // 1. ARRANGE
    const bookingId = "bk-ext-002";
    const dummyBooking = {
      id: bookingId,
      user_id: "user-beta",
      kamar_id: "kamar-102",
      durasi_sewa: 3,
      status: "PENDING",
    };

    const activeEndDate = new Date();
    activeEndDate.setDate(activeEndDate.getDate() + 20);

    const dummyLastActiveBooking = {
      id: "bk-prev",
      user_id: "user-beta",
      status: "APPROVED",
      tanggal_berakhir: activeEndDate,
    };

    (prisma.booking.findUnique as any).mockResolvedValue(dummyBooking);
    (prisma.booking.findFirst as any).mockResolvedValue(dummyLastActiveBooking);
    (prisma.booking.update as any).mockResolvedValue({ id: bookingId, status: "APPROVED" });

    // 2. ACT
    const result = await approveBooking(bookingId);

    // 3. ASSERT
    expect(result).toBeUndefined();
    const updateCallArg = (prisma.booking.update as any).mock.calls[0][0];
    const newCalculatedEndDate = new Date(updateCallArg.data.tanggal_berakhir);

    const expectedEndDate = new Date(activeEndDate);
    expectedEndDate.setMonth(expectedEndDate.getMonth() + 3);
    expect(newCalculatedEndDate.getMonth()).toBe(expectedEndDate.getMonth());
  });

  test("TC-BOOK-04 (Negatif): rejectBooking gagal jika pesanan tidak ditemukan", async () => {
    (prisma.booking.findUnique as any).mockResolvedValue(null);

    const result = await rejectBooking("bk-non-exist");

    expect(result).toEqual({ error: "Pesanan tidak ditemukan" });
    expect(prisma.booking.update).not.toHaveBeenCalled();
  });

  test("TC-BOOK-05 (Positif): rejectBooking berhasil mengubah status pesanan menjadi REJECTED", async () => {
    (prisma.booking.findUnique as any).mockResolvedValue({ id: "bk-rej-1" });
    (prisma.booking.update as any).mockResolvedValue({ id: "bk-rej-1", status: "REJECTED" });

    const result = await rejectBooking("bk-rej-1");

    expect(result).toBeUndefined();
    expect(prisma.booking.update).toHaveBeenCalledWith({
      where: { id: "bk-rej-1" },
      data: { status: "REJECTED" },
    });
    expect(revalidatePath).toHaveBeenCalledWith("/admin/verifikasi");
  });
});
