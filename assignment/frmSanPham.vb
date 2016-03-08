Imports System.Data.SqlClient
Imports System.Data.DataTable

Public Class frmSanPham
    Dim db As New DataTable
    Dim chuoiketnoi As String = "Data Source=PC264\MISASME2012;Integrated Security=True"
    Dim conn As SqlConnection = New SqlConnection(chuoiketnoi)
    Private Sub btnThem_Click(sender As Object, e As EventArgs) Handles btnThem.Click
        reset()
    End Sub

    Private Sub reset()
        txtDongia.Text = ""
        txtMaSP.Text = ""
        txtSoluong.Text = ""
        txtTenSP.Text = ""
        txtMaSP.Focus()
    End Sub
    Private Sub frmSanPham_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadData()
    End Sub

    Private Sub btnLuu_Click(sender As Object, e As EventArgs) Handles btnLuu.Click
        Dim conn As SqlConnection = New SqlConnection(chuoiketnoi)
        Dim query As String = "insert into SANPHAM1 values(@MASP,@TENSP,@SOLUONG,@DONGIA)"
        Dim save As SqlCommand = New SqlCommand(query, conn)
        conn.Open()
        Try
            If txtMaSP.Text = "" Then
                MessageBox.Show("Chua nhap mã sản phẩm")
                txtMaSP.Focus()
            ElseIf txtTenSP.Text = "" Then
                MessageBox.Show("Chua nhap Tên sản phẩm")
                txtTenSP.Focus()
            ElseIf txtSoluong.Text = "" Then
                MessageBox.Show("Chua nhap Số lượng")
                txtSoluong.Focus()
            ElseIf txtDongia.Text = "" Then
                MessageBox.Show("Chua nhap đơn giá")
                txtDongia.Focus()
            Else
                save.Parameters.AddWithValue("@MASP", txtMaSP.Text)
                save.Parameters.AddWithValue("@TENSP", txtTenSP.Text)
                save.Parameters.AddWithValue("@SOLUONG", txtSoluong.Text)
                save.Parameters.AddWithValue("@DONGIA", txtDongia.Text)
                save.ExecuteNonQuery()
                conn.Close()
                MessageBox.Show("Lưu thành công")
                'Sau khi nhập thành công, tự động làm mới các khung textbox, combox và date....
                txtMaSP.Text = Nothing
                txtTenSP.Text = Nothing
                txtSoluong.Text = Nothing
                txtDongia.Text = Nothing
                'LoadData()
            End If
        Catch ex As Exception
            MessageBox.Show("Không được trùng mã sản phẩm", "Lỗi", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error)
        End Try
        'Làm mới lại bảng sau khi lưu thành công
        Dim refesh As SqlDataAdapter = New SqlDataAdapter("select MaSP as 'Mã sản phẩm', TenSP as 'Tên sản phẩm', NgaySX as 'Ngày sản xuất', HangSX as 'Hãng sản xuất' from SANPHAM", conn)
        db.Clear()
        refesh.Fill(db)
        dgvSP.DataSource = db.DefaultView
    End Sub

    Private Sub dgvSP_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvSP.CellContentClick
        Dim click As Integer = dgvSP.CurrentCell.RowIndex
        txtMaSP.Text = dgvSP.Item(0, click).Value
        txtTenSP.Text = dgvSP.Item(1, click).Value
        txtSoluong.Text = dgvSP.Item(2, click).Value
        txtDongia.Text = dgvSP.Item(3, click).Value

    End Sub

    Private Sub LoadData()
        Dim conn As SqlConnection = New SqlConnection(chuoiketnoi)
        Dim refesh As SqlDataAdapter = New SqlDataAdapter("select MaSP as 'Mã SP' ,TenSP as 'Tên Sản phẩm', Soluong as 'Số lượng', Dongia as 'Đơn giá', Soluong * Dongia as 'Thành tiền' from SANPHAM1", conn)
        conn.Open()
        'db.Clear()
        refesh.Fill(db)
        dgvSP.DataSource = db.DefaultView
        'conn.Close()
    End Sub

    Private Sub btnXoa_Click(sender As Object, e As EventArgs) Handles btnXoa.Click
        Dim delquery As String = "delete from SANPHAM1 where MaSP=@MASP"
        Dim delete As SqlCommand = New SqlCommand(delquery, conn)
        Try
            If txtMaSP.Text = "" Then
                MessageBox.Show("Nhap MaSP cần xóa", "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Error)
                txtMaSP.Focus()
            Else
                'Dim resulft As DialogResult = MessageBox.Show("Bạn muốn xóa không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                'If resulft = Windows.Forms.DialogResult.Yes Then
                'conn.Open()
                'delete.Parameters.AddWithValue("@MASP", txtMaSP.Text)
                'delete.ExecuteNonQuery()
                'conn.Close()
                MessageBox.Show("Xóa thành công")
                'Dim sql As String = <sql>
                'delete from sanpham where masp = '{0}'
                ' </sql>
                'Sql = String.Format(sql, txtMaSP.Text)
                'excutenonquery(sql)
                'LoadData()
            End If
        Catch ex As Exception
            MessageBox.Show("Nhập đúng mã sản phẩm cần xóa", "Lỗi", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error)
        End Try
        'làm mới bảng
        db.Clear()
        dgvSP.DataSource = db
        dgvSP.DataSource = Nothing
        LoadData()
    End Sub

    Private Sub btnSua_Click(sender As Object, e As EventArgs) Handles btnSua.Click
        Dim conn As SqlConnection = New SqlConnection(chuoiketnoi)
        Dim query As String = "update SANPHAM1 set TenSP=@TENSP, Soluong=@SOLUONG, Dongia=@DONGIA where MaSP=@MASP"
        Dim save As SqlCommand = New SqlCommand(query, conn)
        conn.Open()
        Try
            txtMaSP.Focus()
            If txtMaSP.Text = "" Then
                MessageBox.Show("Bạn chưa nhập mã sản phẩm", "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Error)
            Else
                txtMaSP.Focus()
                If txtTenSP.Text = "" Then
                    MessageBox.Show("Bạn chưa nhập tên sản phẩm", "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Error)
                Else
                    txtTenSP.Focus()
                    If txtSoluong.Text = "" Then
                        MessageBox.Show("Bạn chưa nhập ngày sản xuất", "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Error)
                    Else
                        txtSoluong.Focus()
                        If txtDongia.Text = "" Then
                            MessageBox.Show("Bạn chưa nhập hãng sản xuất", "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Error)
                        Else
                            Dim sql As String = <sql>
                             update sanpham
                            set  TenSP=N'{0}', NgaySX= '{1}', HangSX= '{2}' 
                            where MaSP='{3}' 
                             </sql>
                            sql = String.Format(sql, txtTenSP.Text, txtSoluong.Text, txtDongia.Text, txtMaSP.Text)
                            excutenonquery(sql)


                            MessageBox.Show("Cập nhật thành công")
                            txtMaSP.Text = Nothing
                            txtTenSP.Text = Nothing
                            txtSoluong.Text = Nothing
                            txtDongia.Text = Nothing
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Không thành công", "Lỗi", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error)
        End Try

        'Sau khi cập nhật xong sẽ tự làm mới lại bảng
        db.Clear()
        dgvSP.DataSource = db
        dgvSP.DataSource = Nothing
        LoadData()
        'If btnSua.Text = "Sửa" Then
        'txtMaSP.ReadOnly = True
        'btnSua.Text = "Update"
        'txtTenSP.Focus()
        'ElseIf btnSua.Text = "Update" Then

        'save.Parameters.AddWithValue("@MASP", txtMaSP.Text)
        'save.Parameters.AddWithValue("@TENSP", txtTenSP.Text)
        'save.Parameters.AddWithValue("@SOLUONG", txtSoluong.Text)
        'save.Parameters.AddWithValue("@DONGIA", txtDongia.Text)
        'save.ExecuteNonQuery()
        'conn.Close()
        'MessageBox.Show("Update thành công")
        'txtMaSP.ReadOnly = False
        'btnSua.Text = "Sửa"
        'LoadData()
        'End If
    End Sub
End Class