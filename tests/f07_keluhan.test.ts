import { updateTiketStatus } from "@/app/admin/keluhan/actions";
import prisma from "@/lib/prisma";
import { revalidatePath } from "next/cache";

jest.mock("@/lib/prisma", () => ({
  __esModule: true,
  default: {
    tiket: {
      update: jest.fn(),
    },
  },
}));

jest.mock("next/cache", () => ({
  revalidatePath: jest.fn(),
}));

describe("F-07: Kelola Tiket Keluhan - Unit Testing (updateTiketStatus)", () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test("TC-TKT-01 (Positif): berhasil mengubah status tiket keluhan menjadi PROSES", async () => {
    // 1. ARRANGE
    const tiketId = "tiket-001";
    const targetStatus = "PROSES" as any;
    (prisma.tiket.update as any).mockResolvedValue({ id: tiketId, status: targetStatus });

    // 2. ACT
    await updateTiketStatus(tiketId, targetStatus);

    // 3. ASSERT
    expect(prisma.tiket.update).toHaveBeenCalledWith({
      where: { id: tiketId },
      data: { status: targetStatus },
    });
    expect(revalidatePath).toHaveBeenCalledWith("/admin/keluhan");
  });

  test("TC-TKT-02 (Positif): berhasil mengubah status tiket keluhan menjadi SELESAI", async () => {
    // 1. ARRANGE
    const tiketId = "tiket-002";
    const targetStatus = "SELESAI" as any;
    (prisma.tiket.update as any).mockResolvedValue({ id: tiketId, status: targetStatus });

    // 2. ACT
    await updateTiketStatus(tiketId, targetStatus);

    // 3. ASSERT
    expect(prisma.tiket.update).toHaveBeenCalledWith({
      where: { id: tiketId },
      data: { status: targetStatus },
    });
    expect(revalidatePath).toHaveBeenCalledWith("/admin/keluhan");
  });
});
