import io
import re
import streamlit as st
import openpyxl
from google import genai


# =====================================================
# CẤU HÌNH TRANG WEB
# =====================================================

st.set_page_config(
    page_title="AI Smart Eval THCS",
    page_icon="📝",
    layout="centered"
)

st.title("📝 CHẤM BÀI THỰC HÀNH TIN HỌC")
st.caption("Hệ thống hỗ trợ chấm bài thực hành Excel")


# =====================================================
# HÀM KIỂM TRA CÔNG THỨC
# =====================================================

def co_ham(cong_thuc, ten_ham):
    if not isinstance(cong_thuc, str):
        return False

    if not cong_thuc.startswith("="):
        return False

    mau = rf"\b{re.escape(ten_ham)}\s*\("

    return re.search(
        mau,
        cong_thuc,
        re.IGNORECASE
    ) is not None


# =====================================================
# HÀM KIỂM TRA VIỀN
# =====================================================

def co_du_vien(cell):
    border = cell.border

    if border is None:
        return False

    cac_canh = [
        border.left,
        border.right,
        border.top,
        border.bottom
    ]

    return all(
        canh is not None
        and canh.style is not None
        for canh in cac_canh
    )


# =====================================================
# HÀM CHẤM BÀI EXCEL
# =====================================================

def cham_excel(file_bytes):

    wb = openpyxl.load_workbook(
        io.BytesIO(file_bytes),
        data_only=False
    )

    sheet = wb.active

    # ---------------------------------
    # Xác định các hàng có dữ liệu
    # ---------------------------------

    cac_hang = []

    for r in range(2, sheet.max_row + 1):

        if sheet[f"A{r}"].value not in (None, ""):
            cac_hang.append(r)

    if not cac_hang:
        return None

    ket_qua = []

    # =================================================
    # TIÊU CHÍ 1
    # Cột F dùng hàm AVERAGE - 4 điểm
    # =================================================

    dung_f = []
    sai_f = []

    for r in cac_hang:

        dia_chi = f"F{r}"

        if co_ham(
            sheet[dia_chi].value,
            "AVERAGE"
        ):
            dung_f.append(dia_chi)
        else:
            sai_f.append(dia_chi)

    diem_f = (
        len(dung_f)
        / len(cac_hang)
        * 4
    )

    ket_qua.append({
        "ten": "Cột F sử dụng hàm AVERAGE",
        "diem": round(diem_f, 2),
        "diem_max": 4,
        "sai": sai_f
    })


    # =================================================
    # TIÊU CHÍ 2
    # Cột G dùng hàm IF - 4 điểm
    # =================================================

    dung_g = []
    sai_g = []

    for r in cac_hang:

        dia_chi = f"G{r}"

        if co_ham(
            sheet[dia_chi].value,
            "IF"
        ):
            dung_g.append(dia_chi)
        else:
            sai_g.append(dia_chi)

    diem_g = (
        len(dung_g)
        / len(cac_hang)
        * 4
    )

    ket_qua.append({
        "ten": "Cột G sử dụng hàm IF",
        "diem": round(diem_g, 2),
        "diem_max": 4,
        "sai": sai_g
    })


    # =================================================
    # TIÊU CHÍ 3
    # Hàng tiêu đề A1:G1 in đậm - 1 điểm
    # =================================================

    dung_dam = 0
    sai_dam = []

    for cot in range(1, 8):

        cell = sheet.cell(
            row=1,
            column=cot
        )

        if bool(cell.font.bold):
            dung_dam += 1
        else:
            sai_dam.append(
                cell.coordinate
            )

    diem_dam = (
        dung_dam / 7
    )

    ket_qua.append({
        "ten": "Hàng tiêu đề in đậm",
        "diem": round(diem_dam, 2),
        "diem_max": 1,
        "sai": sai_dam
    })


    # =================================================
    # TIÊU CHÍ 4
    # Bảng A1:G... có viền - 1 điểm
    # =================================================

    hang_cuoi = max(cac_hang)

    tong_o = 0
    so_o_dung = 0
    sai_vien = []

    for r in range(1, hang_cuoi + 1):

        for c in range(1, 8):

            cell = sheet.cell(
                row=r,
                column=c
            )

            tong_o += 1

            if co_du_vien(cell):
                so_o_dung += 1
            else:
                sai_vien.append(
                    cell.coordinate
                )

    diem_vien = (
        so_o_dung
        / tong_o
    )

    ket_qua.append({
        "ten": "Bảng được kẻ viền",
        "diem": round(diem_vien, 2),
        "diem_max": 1,
        "sai": sai_vien[:20]
    })


    # =================================================
    # TỔNG ĐIỂM
    # =================================================

    tong_diem = round(
        sum(
            tc["diem"]
            for tc in ket_qua
        ),
        2
    )

    return {
        "tong_diem": tong_diem,
        "chi_tiet": ket_qua
    }


# =====================================================
# GIAO DIỆN HỌC SINH
# =====================================================

st.subheader("🎓 Nộp bài")

ho_ten = st.text_input(
    "Họ và tên học sinh"
)

lop = st.text_input(
    "Lớp",
    placeholder="Ví dụ: 7A1"
)

file_nop = st.file_uploader(
    "Chọn bài thực hành Excel",
    type=["xlsx"]
)


# =====================================================
# CHẤM BÀI
# =====================================================

if st.button(
    "🚀 Nộp bài và chấm điểm",
    type="primary"
):

    if not ho_ten.strip():

        st.warning(
            "Em hãy nhập họ và tên."
        )

    elif not lop.strip():

        st.warning(
            "Em hãy nhập lớp."
        )

    elif file_nop is None:

        st.warning(
            "Em hãy chọn file bài làm."
        )

    else:

        try:

            file_bytes = (
                file_nop.getvalue()
            )

            ket_qua = cham_excel(
                file_bytes
            )

            if ket_qua is None:

                st.error(
                    "Không tìm thấy dữ liệu "
                    "để chấm."
                )

            else:

                st.success(
                    "✅ Đã chấm xong!"
                )

                st.metric(
                    "🎯 Điểm",
                    f"{ket_qua['tong_diem']}/10"
                )

                st.write(
                    f"**Học sinh:** {ho_ten}"
                )

                st.write(
                    f"**Lớp:** {lop}"
                )

                st.subheader(
                    "📊 Chi tiết"
                )

                for tc in ket_qua[
                    "chi_tiet"
                ]:

                    if len(tc["sai"]) == 0:
                        icon = "✅"
                    else:
                        icon = "⚠️"

                    st.write(
                        f"{icon} "
                        f"**{tc['ten']}**: "
                        f"{tc['diem']}/"
                        f"{tc['diem_max']} điểm"
                    )

                    if tc["sai"]:

                        st.caption(
                            "Cần xem lại: "
                            + ", ".join(
                                tc["sai"]
                            )
                        )


                # ================================
                # NHẬN XÉT BẰNG GEMINI
                # ================================

                try:

                    api_key = (
                        st.secrets[
                            "GEMINI_API_KEY"
                        ]
                    )

                    client = genai.Client(
                        api_key=api_key
                    )

                    prompt = f"""
Bạn là giáo viên Tin học THCS.

Học sinh:
{ho_ten}

Lớp:
{lop}

Kết quả chấm do Python xác định:
{ket_qua}

Điểm chính thức:
{ket_qua['tong_diem']}/10

QUY TẮC:

1. Không được thay đổi điểm.
2. Không giải hộ học sinh.
3. Không viết sẵn công thức hoàn chỉnh.
4. Chỉ rõ vị trí học sinh cần xem lại.
5. Đặt câu hỏi gợi mở để học sinh tự sửa.
6. Giọng văn thân thiện, khích lệ.
7. Nhận xét ngắn gọn, phù hợp học sinh THCS.

Trình bày:

🌟 Điểm làm tốt

🔍 Nội dung cần xem lại

💡 Câu hỏi gợi ý
"""

                    response = (
                        client.models
                        .generate_content(
                            model=(
                                "gemini-3.8-flash"
                            ),
                            contents=prompt
                        )
                    )

                    st.subheader(
                        "🤖 Nhận xét của AI"
                    )

                    st.markdown(
                        response.text
                    )

                except Exception:

                    st.info(
                        "Điểm đã được chấm. "
                        "Hiện chưa tạo được "
                        "nhận xét AI."
                    )

        except Exception as e:

            st.error(
                "Không đọc được bài làm: "
                + str(e)
            )
