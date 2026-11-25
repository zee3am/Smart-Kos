import { hapusKamar } from "@/app/admin/kamar/actions";
import prisma from "@/lib/prisma";
import { revalidatePath } from "next/cache";

jest.mock("@/lib/prisma", () => ({
  __esModule: true,
  default: {
    kamar: {
      delete: jest.fn(),
    },
  },
}));

jest.mock("next/cache", () => ({
  revalidatePath: jest.fn(),
}));

describe("F-04: Kelola Data Kamar - Unit Testing (hapusKamar)", () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test("TC-R00M-01 (Positif): berhasil menghapus data kamar dan merevalidasi cache halaman /admin/kamar", async () => {
    // 1. ARRANGE
    const kamarId = "kamar-uuid-101";
    // STUB: return deleted object
    (prisma.kamar.delete as any).mockResolvedValue({ id: kamarId });

    // 2. ACT
    await hapusKamar(kamarId);

    // 3. ASSERT (MOCK Verification)
    expect(prisma.kamar.delete).toHaveBeenCalledWith({
      where: { id: kamarId },
    });
    expect(revalidatePath).toHaveBeenCalledWith("/admin/kamar");
  });
});
