import os
import pandas as pd
from flask import Flask, render_template, request, redirect, send_file
from werkzeug.utils import secure_filename
from io import BytesIO

app = Flask(__name__)

# Định nghĩa thư mục lưu trữ file tải lên
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def calculate_cost_and_commission(quantity, product_type, selling_price,
                                  weight_per_bag=25,discount_amount=0):

    # Chuyển đổi giá bán thực tế sang kiểu số (float)
    selling_price = float(selling_price)

    # Chuyển `product_type` thành chữ thường để không phân biệt chữ hoa hay chữ thường
    product_type = product_type.lower()

    # Định nghĩa giá cơ bản và hoa hồng dựa trên loại sản phẩm
    if product_type == "keo dán gạch bricon thông dụng - màu trắng":
        base_price_per_bag = 250000
        commission_per_bag = 15000
        price_brackets = [(10, 19, 165000), (20, 49, 160000), (50, 99, 155000), (100, 159, 140000), (160, float('inf'), 135000)]
    elif product_type == "keo dán gạch bricon thông dụng - màu xám":
        base_price_per_bag = 230000
        commission_per_bag = 15000
        price_brackets = [(10, 19, 150000), (20, 49, 145000), (50, 99, 140000), (100, 159, 125000), (160, float('inf'), 120000)]
    elif product_type == "keo dán gạch bricon extra - màu trắng":
        base_price_per_bag = 350000
        commission_per_bag = 15000
        price_brackets = [(10, 19, 225000), (20, 49, 220000), (50, 99, 210000), (100, 159, 200000), (160, float('inf'), 195000)]
    elif product_type == "keo dán gạch bricon extra - màu xám":
        base_price_per_bag = 295000
        commission_per_bag = 15000
        price_brackets = [(10, 19, 195000), (20, 49, 190000), (50, 99, 180000), (100, 159, 170000), (160, float('inf'), 160000)]
    elif product_type == "keo chà ron hộp 30kg/thùng":
        base_price_per_bag = 810000
        commission_per_bag = 25000
        price_brackets = [(5, 9, 375000), (10, 29, 345000), (30, float('inf'), 285000)]
        weight_per_bag = 30  
    elif product_type == "keo chà ron 24 túi/thùng":
        base_price_per_bag = 600000
        commission_per_bag = 30000
        price_brackets = [(5, 9, 290000), (10, 39, 265000), (40, float('inf'), 230000)]
        weight_per_bag = 24  
    else:
        raise ValueError("Loại sản phẩm không hợp lệ")
    
    # Xác định giá đề xuất dựa trên số lượng mua
    price_per_bag = base_price_per_bag
    for min_qty, max_qty, price in price_brackets:
        if min_qty <= quantity <= max_qty:
            price_per_bag = price
            break

    # Tính tổng giá sản phẩm phải thu (giá đề xuất)
    total_must_collect = price_per_bag * quantity

    # Tính tổng chi phí thực thu (giá bán thực tế)
    total_actual_collect = selling_price * quantity

    # Tính hoa hồng cơ bản
    basic_commission = commission_per_bag * quantity

    # Tính tổng khoản chênh lệch và cộng vào hoa hồng
    difference_per_bag = selling_price - price_per_bag
    total_difference = difference_per_bag * quantity

    # Tổng hoa hồng bao gồm cả chênh lệch
    total_commission = basic_commission + total_difference

    # Trừ số tiền tặng khách từ tổng hoa hồng
    final_commission = total_commission - discount_amount

    # Tính phí vận chuyển
    total_weight = quantity * weight_per_bag
    if (product_type == "keo chà ron hộp 30kg/thùng" or product_type == "keo chà ron 24 túi/thùng") and quantity >= 5:
        shipping_fee = 0  # Miễn phí nếu mua từ 5 thùng trở lên
    elif quantity < 10:
        if total_weight <= 25:
            shipping_fee = 38450
        elif total_weight <= 100:
            shipping_fee = 130000
        else:
            shipping_fee = 230000
    else:
        shipping_fee = 0  # Miễn phí nếu mua từ 10 bao trở lên


    return {
        "Tổng chi phí phải thu(Giá đề xuất)": f"{total_must_collect:,.0f}₫",
        "Tổng chi phí thực thu(Giá bán thực tế)": f"{total_actual_collect:,.0f}₫",
        "Phí Ship(Đàm phán)": f"{shipping_fee:,.0f}₫",
        "Hoa hồng cơ bản(Tiền chiết khấu)": f"{basic_commission:,.0f}₫",
        "Tiền chênh lệch": f"{total_difference:,.0f}₫",
        "Tổng hoa hồng(Lợi nhuận)": f"{final_commission:,.0f}₫"
    }

@app.route('/', methods=['GET', 'POST'])
def input_form():
    if request.method == 'POST':
        quantities = request.form.getlist('quantity[]')
        product_types = request.form.getlist('product_type[]')
        selling_prices = request.form.getlist('selling_price[]')
        total_amounts = request.form.getlist('total_amount[]')
        discount_amount = float(request.form.get('discount_amount', 0))

        #Kết quả
        results = []
        # Danh sách lưu các thông tin sản phẩm nhập vào
        products_info = []
        
        for i in range(len(quantities)):
            quantity = int(quantities[i])
            product_type = product_types[i]
            selling_price = selling_prices[i]
            total_amount = total_amounts[i]

            # Xử lý trường hợp không có `selling_price`, dùng `total_amount`
            if selling_price:
                selling_price = float(selling_price)
            elif total_amount:
                selling_price = float(total_amount) / quantity
            else:
                return render_template('input_form.html', error="Vui lòng nhập giá bán hoặc tổng tiền thu về.")

            # Lưu thông tin sản phẩm vào danh sách
            product_info = {
                'quantity': quantity,
                'product_type': product_type,
                'selling_price': selling_price,
                'total_amount': total_amount if total_amount else selling_price * quantity
            }
            products_info.append(product_info)


            # Tính toán cho từng sản phẩm và lưu kết quả vào danh sách
            result = calculate_cost_and_commission(quantity, product_type, selling_price, discount_amount=discount_amount)
            results.append(result)

        return render_template('input_form.html', results=results, products_info=products_info, discount_amount=discount_amount)

    return render_template('input_form.html')


@app.route('/upload', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            return "Không có file nào trong yêu cầu"
        
        file = request.files['file']
        if file.filename == '':
            return "Không có file nào được chọn để tải lên"

        if file:
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)

            # Đọc file Excel và chuẩn hóa tên cột thành chữ thường
            df = pd.read_excel(file_path)
            df.columns = df.columns.str.lower()
            results = []

            for _, row in df.iterrows():
                quantity = row['số lượng']
                product_type = row['sản phẩm']
                discount_amount = row.get('số tiền tặng khách', 0)  # Lấy `discount_amount`, mặc định là 0 nếu thiếu

                selling_price = row.get('giá bán thực tế mỗi sản phẩm')
                total_amount = row.get('tổng chi phí thực thu')

                # Tính `selling_price` từ `total_amount` nếu `selling_price` không có
                if pd.isna(selling_price) and not pd.isna(total_amount):
                    selling_price = total_amount / quantity

                # Tính toán kết quả cho từng dòng
                result = calculate_cost_and_commission(quantity, product_type, selling_price, discount_amount=discount_amount)
                results.append(result)

            # Gộp kết quả với dữ liệu ban đầu và xuất thành file Excel mới
            results_df = pd.DataFrame(results)
            output_df = pd.concat([df, results_df], axis=1)

            output_stream = BytesIO()
            output_df.to_excel(output_stream, index=False, engine='openpyxl')
            output_stream.seek(0)

            return send_file(output_stream, as_attachment=True, download_name="ket_qua.xlsx", mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    return render_template('index.html')


if __name__ == '__main__':
    app.run(debug=True)