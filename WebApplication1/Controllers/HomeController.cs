using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace WebApplication1.Controllers
{
    public class HomeController : Controller
    {
        // GET: Home
        public ActionResult Index()
        {
            XWPFDocument doc = new XWPFDocument();



            #region phanDau
            XWPFParagraph para = doc.CreateParagraph();
            XWPFRun run = para.CreateRun();
            using (FileStream picFile = new FileStream(@"K:\\ThucTap1\\WebApplication1\\WebApplication1\\pic\\p1.jpg", FileMode.Open, FileAccess.Read))
            {
                run.AddPicture(picFile, (int)PictureType.PNG, "p1", 1000 * 2000, 1000 * 600);


            }
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText(" \t\t\t\tCỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM");
            run.IsBold = true;

            //==============
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.SetText("CÔNG TY CP CÔNG NGHỆ VSHARE");
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);

            run = para.CreateRun();
            run.SetText(" \t\t\t\t\t\tĐộc lập – Tự do – Hạnh phúc");
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);


            //===========
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.SetText(" Số:" ); 
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);

            run = para.CreateRun();
            run.FontSize = 10;
            run.IsBold = true;
            run.SetFontFamily("Arial", FontCharRange.None);
            run.SetText(" {{Số ID hợp đồng}}");

            run = para.CreateRun();
            run.SetText("/Vshare");
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);

            //==================
            para = doc.CreateParagraph();
            run = para.CreateRun();
            para = doc.CreateParagraph();
            para.Alignment = ParagraphAlignment.CENTER;
            run = para.CreateRun();
            run.SetText("HỢP ĐỒNG CHO THUÊ XE");
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.FontSize = 14;
            run.IsBold=true;
            para = doc.CreateParagraph();
            run = para.CreateRun();

            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\t\t\t-Căn cứ Bộ Luật dân sự số 91/2015/QH13 nước CHXHCN Việt Nam ngày 01/01/2017");

            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\t\t\t- Căn cứ Luật thương mại số 36/2005/QH11 nước CHXHCN Việt Nam ngày 26/06/2005");

            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\t\t\t- Căn cứ vào khả năng cung cấp và nhu cầu của hai bên.");
            para = doc.CreateParagraph();
            run = para.CreateRun();

            para = doc.CreateParagraph();
            para.Alignment = ParagraphAlignment.LEFT;
            para.SpacingLineRule=LineSpacingRule.AUTO;
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman",FontCharRange.None);
            run.SetText("Hôm nay, ngày {{Ngày thẩm định}}, chúng tôi gồm:");
            #endregion

            #region Doan1
            //1
            para = doc.CreateParagraph();
            para.Alignment = ParagraphAlignment.LEFT;
            run = para.CreateRun();
            run.FontSize = 11;
            run.IsBold=true;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("BÊN CHO THUÊ XE (BÊN A): CÔNG TY CỔ PHẦN CÔNG NGHỆ VSHARE ");

            
            //2
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText(" \tĐịa chỉ:		Park 2, Times City, 458 Minh Khai, Hai Bà Trưng, Hà Nội.");

            //run = para.CreateRun();
            //using (FileStream picFile = new FileStream(@"K:\\ThucTap1\\WebApplication1\\WebApplication1\\pic\\p2.jpg", FileMode.Open, FileAccess.Read))
            //{
                
            //    run.AddPicture(picFile, (int)PictureType.PNG, "p2", 700 * 1800, 900 * 600);
                

            //}
            //3
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\tMST:		0110032942 ");
            //4
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\tSTK: 	  	19035651301016 tại NH Kỹ Thương VN (Techcombank), CN Hai Bà Trưng");

            //5
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.Ascii);
            run.SetText("\tĐại diện:");

            run = para.CreateRun();
            run.FontSize = 10;
            run.IsBold = true;
            run.SetFontFamily("Arial", FontCharRange.Ascii);
            run.SetText("{ { Người đại diện làm hợp đồng } }");


            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.Ascii);
            run.SetText(" \tĐiện thoại:");

            run = para.CreateRun();
            run.FontSize = 10;
            run.IsBold = true;
            run.SetFontFamily("Arial", FontCharRange.Ascii);
            run.SetText("{{Số điện thoại người làm hợp đồng}}");
            #endregion

            #region Doan2
            //1
            para = doc.CreateParagraph();
            para.Alignment = ParagraphAlignment.LEFT;
            run = para.CreateRun();
            run.FontSize = 11;
            run.IsBold = true;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("BÊN THUÊ XE (BÊN B): ");

            run = para.CreateRun();
            run.FontSize = 10;
            run.IsBold = true;
            run.SetFontFamily("Arial", FontCharRange.Ascii);
            run.SetText("{{Họ tên khách hàng}} ");
            
            //2
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\tĐịa chỉ hiện tại:");

            run = para.CreateRun();
            run.FontSize = 10;
            run.IsBold = true;
            run.SetFontFamily("Arial", FontCharRange.Ascii);
            run.SetText("{{Địa chỉ hiện tại khách hàng}}");

            //3
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\tSố điện thoại: ");

            run = para.CreateRun();
            run.FontSize = 10;
            run.IsBold = true;
            run.SetFontFamily("Arial", FontCharRange.None);
            run.SetText("{{Số điện thoại khách hàng}}");
            
            //4
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\tCCCD/ CMND số:");

            run = para.CreateRun();
            run.FontSize = 10;
            run.IsBold = true;
            run.SetFontFamily("Arial", FontCharRange.Ascii);
            run.SetText("{{Số CCCD khách hàng}}");

            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\t\tCấp ngày:");

            run = para.CreateRun();
            run.FontSize = 10;
            run.IsBold = true;
            run.SetFontFamily("Arial", FontCharRange.Ascii);
            run.SetText("{{Ngày cấp CCCD khách hàng}}");

            //5
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\tGPLX số:");

            run = para.CreateRun();
            run.FontSize = 10;
            run.IsBold = true;
            run.SetFontFamily("Arial", FontCharRange.Ascii);
            run.SetText("\t{{Số GPLX khách hàng}}");


            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\t\t\t\tCấp ngày: ");

            run = para.CreateRun();
            run.FontSize = 10;
            run.IsBold = true;
            run.SetFontFamily("Arial", FontCharRange.Ascii);
            run.SetText("{{Ngày cấp GPLX khách hàng}}");

            //6
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\tSau khi bàn bạc, thỏa thuận, hai bên cùng nhất trí ký hợp đồng thuê xe với các điều khoản sau:");
            run.IsBold = true;
            run.IsItalic = true;
            #endregion

            #region Doan3
            //1
            para = doc.CreateParagraph();
            para.Alignment = ParagraphAlignment.LEFT;
            run = para.CreateRun();
            run.FontSize = 11;
            run.IsBold = true;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("Điều I: Nội dung hợp đồng ");

            //2
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\tBên A đồng ý cho Bên B thuê xác xe ô tô để phục vụ mục đích đi lại:");

            //3
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\tHiệu xe:");

            run = para.CreateRun();
            run.FontSize = 10;
            run.IsBold = true;
            run.SetFontFamily("Arial", FontCharRange.Ascii);
            run.SetText("{{Nhãn hiệu xe cho thuê}}");

            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\tBiển số:");

            run = para.CreateRun();
            run.FontSize = 10;
            run.IsBold = true;
            run.SetFontFamily("Arial", FontCharRange.Ascii);
            run.SetText("{{Biển số xe cho thuê}}");

            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\tMàu:");

            run = para.CreateRun();
            run.FontSize = 10;
            run.IsBold = true;
            run.SetFontFamily("Arial", FontCharRange.Ascii);
            run.SetText("{{Màu xe cho thuê}}");

            //4
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\tThời gian thuê xe:");

            run = para.CreateRun();
            run.FontSize = 10;
            run.IsBold = true;
            run.SetFontFamily("Arial", FontCharRange.Ascii);
            run.SetText("\t{{Thời lượng xe cho thuê (ngày) }}");

            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("ngày");

            //5
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\tBắt đầu từ:");

            run = para.CreateRun();
            run.FontSize = 10;
            run.IsBold = true;
            run.SetFontFamily("Arial", FontCharRange.Ascii);
            run.SetText("{{Thời gian dự kiến bắt đầu nhận xe ( giờ/ngày/tháng/năm) }} ");

            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\tĐến:");

            run = para.CreateRun();
            run.FontSize = 10;
            run.IsBold = true;
            run.SetFontFamily("Arial", FontCharRange.Ascii);
            run.SetText("{{Thời gian dự \tkiến trả xe ( giờ/ngày/tháng/năm)}} ");
            #endregion

            #region Doan4
            //1
            para = doc.CreateParagraph();
            para.Alignment = ParagraphAlignment.LEFT;
            run = para.CreateRun();
            run.FontSize = 11;
            run.IsBold = true;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("Điều II: Thanh toán  ");

            run = para.CreateRun();
            run.FontSize = 10;
            run.IsItalic = true;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("(giá chưa bao gồm VAT)");

            //2
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\tĐơn giá sau khuyến mãi:");

            run = para.CreateRun();
            run.FontSize = 10;
            run.IsBold = true;
            run.SetFontFamily("Arial", FontCharRange.Ascii);
            run.SetText("{{Giá tiền thuê một ngày}} / ");

            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("ngày");

            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\t\t\tGiới hạn hành trình: ");

            run = para.CreateRun();
            run.FontSize = 10;
            run.IsBold = true;
            run.SetFontFamily("Arial", FontCharRange.Ascii);
            run.SetText("{{Tổng giới hạn \tkm xe}} ");

            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("Km");

            //3
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\tĐặt cọc giữ xe:");

            run = para.CreateRun();
            run.FontSize = 10;
            run.IsBold = true;
            run.SetFontFamily("Arial", FontCharRange.Ascii);
            run.SetText("{{Tiền cọc}}{{Tài sản cọc}}");

 
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\t\t\tPhí phát sinh Km:");

            run = para.CreateRun();
            run.FontSize = 10;
            run.IsBold = true;
            run.SetFontFamily("Arial", FontCharRange.Ascii);
            run.SetText("{{Phí vượt/1km}}");

            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("/Km");

            //4
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\tTổng tiền thuê xe:");

            run = para.CreateRun();
            run.FontSize = 10;
            run.IsBold = true;
            run.SetFontFamily("Arial", FontCharRange.Ascii);
            run.SetText("{{Tổng tiền thuê}}");

            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\t\t\tPhí phát sinh thời gian: 100.000 đ/giờ ");

            //5
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\tThanh toán khi nhận xe:");

            run = para.CreateRun();
            run.FontSize = 10;
            run.IsBold = true;
            run.SetFontFamily("Arial", FontCharRange.Ascii);
            run.SetText("{{Số Tiền Cần Thanh Toán}}");

            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\t\t\tPhí giao nhận xe:");

            run = para.CreateRun();
            run.FontSize = 10;
            run.IsBold = true;
            run.SetFontFamily("Arial", FontCharRange.Ascii);
            run.SetText("{{Phí giao nhận \txe}}");
            #endregion

            #region Doan5
            //1
            para = doc.CreateParagraph();
            para.Alignment = ParagraphAlignment.LEFT;
            run = para.CreateRun();
            run.FontSize = 11;
            run.IsBold = true;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("Điều III: Quyền hạn trách nhiệm của Bên A ");

            //2
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("1. Bên A có trách nhiệm giao xe cho bên B đúng chủng loại," +
                " hoạt động bình thường, địa điểm đã thỏa thuận và cung cấp thông tin cần " +
                "thiết để sử dụng phương tiện. Trường hợp bất khả kháng, xe xảy ra sự cố thì" +
                " bên A sẽ thay thế xe có giá trị tương đương để bên B sử dụng.");

            //3
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("2. Cung cấp đầy đủ các giấy tờ bao gồm: Đăng ký xe (1), Kiểm định xe (2), Bảo hiểm xe (3)");

            //4
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("3. Bên A có quyền yêu cầu bên B trả xe trước thời hạn hợp đồng nếu bên A" +
                " phát hiện bên B sử dụng sai mục đích hoặc bên B vi phạm các điều khoản đã thỏa thuận trong hợp đồng này.");

            //5
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("4. Bên A có quyền sử dụng tài sản thế chấp của bên B để thanh " +
                "toán hợp đồng (trong trường hợp bên B thực hiện không đúng hoặc không đủ nghĩa vụ của " +
                "hợp đồng với bên A) và các chi phí phát sinh khác. Nếu giá trị tài sản thế chấp của bên B " +
                "thấp hơn các chi phí phát sinh thì bên B có trách nhiệm thanh toán thêm bằng tiền mặt.");


            //6
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("5. Bên A có quyền đơn phương chấm dứt hợp đồng nếu bên B (hoặc tài xế) " +
                "sử dụng xe không thành thạo, hoặc thành thạo nhưng không đảm bảo an toàn như : " +
                "say rượu, sử dụng chất kích thích... Mọi chi phí bên thuê xe vẫn phải thanh toán trong trường hợp này.");

            //7
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("6. Trong thời gian thuê xe, nếu bên A không liên lạc được với bên B (trong vòng 12 giờ), bên A " +
                "được quyền nhờ cơ quan chức năng tìm kiếm và thu hồi xe về, bên B phải chịu chi phí thiệt hại này.");

            //8
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("7. Khi bên A phát hiện bên B chạy quá tốc độ bằng phương pháp nghiệp vụ như theo dõi" +
                " trên thiết bị định vị GPS (lắp trên xe) hoặc do phía cơ quan chức năng cung cấp. Bên B có thể đơn phương " +
                "chấm dứt hợp đồng luôn với bên A tại thời điểm đó, toàn bộ chi phí cọc và tiền phát sinh bên B sẽ chịu toàn bộ chi phí.");

            //9
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("8. Nếu bên B bị phạt nguội mà trước hoặc sau khi kết thúc hợp đồng bên A vẫn có quyền phạt bên B." +
                " Mọi chi phí phát sinh như đi lại, ăn uống, nhà nghỉ, phí phạt, bên B sẽ thanh toán toàn bộ cho bên A.");

            #endregion

            #region Doan6
            //1
            para = doc.CreateParagraph();
            para.Alignment = ParagraphAlignment.LEFT;
            run = para.CreateRun();
            run.FontSize = 11;
            run.IsBold = true;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("Điều IV: Quyền hạn và trách nhiệm bên B ");

            //2
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("1. Khi xảy ra sự cố và va quệt làm hư hỏng xe, bên B " +
                "phải thông báo và bồi thường theo hiện trạng cũ mà nơi sửa chữa do " +
                "bên A quy định. Khi sửa chữa, bên B có trách nhiệm bồi thường thiệt hại về " +
                "tiền cho bên A khi xe không lưu hành được theo đơn giá của ngày thuê xe." +
                " (Số ngày xe sửa chữa x Đơn giá thuê) ");

            //3
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("2. Nếu Bên B trả xe mà không có hoặc làm mất giấy tờ xe " +
                "(được quy định tại điều III, khoản 2) hoặc vi phạm luật an toàn giao " +
                "thông đường bộ dẫn đến bị thu xe thì Bên B vẫn phải thanh toán tiền thuê " +
                "xe bình thường cho đến khi trao trả giấy tờ xe và xe đầy đủ.");

            //4
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("3. Nếu trong trường hợp bên B làm mất xe thì bên B phải " +
                "chịu bồi thường 100% giá trị ban đầu của xe cho bên A.");

            //5
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("4. Trong quá trình sử dụng xe bên B chịu hoàn toàn trách nhiệm dân" +
                " sự, hình sự, và luật lệ an toàn giao thông trước pháp luật nếu có phát sinh " +
                "bất cứ chi phí phạt nào tại thời điểm bên B thuê xe, bên B vẫn chịu chi phí đó " +
                "mặc dù hợp đồng đã thanh lý. Bên B tuân thủ đi đúng hành trình đã cam kết với bên A," +
                " nếu có thay đổi phải báo cho bên A biết. Nếu không bên A có quyền đơn phương chấm dứt " +
                "hợp đồng, lấy xe về trước thời hạn.");


            //6
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("5. Hết hạn hợp đồng bên B trả xe ngay cho bên A (như tình trạng xe khi bàn giao)." +
                " Thời gian Bên A nhận xe từ Bên B không muộn hơn giờ quy định ở trên, nếu trả xe sau giờ " +
                "quy định Bên A sẽ tính phí phát sinh : 100.000 vnđ /1 giờ. Trường hợp Bên B trả xe sau 22h00" +
                " bên A sẽ tính chi phí phát sinh là 1 ngày.");

            //7
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("6. Trường hợp xe phát sinh : Khi phát sinh trả trước hạn hợp đồng bên B phải báo bên " +
                "A trước 24h. Trường hợp Bên B không báo trước, bên A sẽ thu phí theo giá trị Hợp Đồng và không hoàn" +
                " trả tiền thừa. Khi phát sinh trả sau hạn hợp đồng, Bên B đi thêm, Bên B phải báo bên A trước 8 tiếng" +
                " so với thời gian hết hạn. Trường hợp Bên B báo muộn sau 8 tiếng chi phí phát sinh cho ngày mới tăng" +
                " 30% giá thuê.");

            //8
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("7. Bên B không giao xe cho người khác sử dụng dưới bất kì hình thức nào hoặc chuyên chở " +
                "các loại vũ khí, chất cháy nổ, hàng quốc cấm cũng như các đồ hải sản, đồ ăn, nước chấm, mắm hoặc" +
                " hàng nặng mùi. Nếu vi phạm sẽ bị phạt từ 1.000.000 vnđ đến 5.000.000 vnđ cũng như toàn bộ chi phí " +
                "khắc phục và giá trị ngày xe không khai thác kinh doanh được.");

            //9
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("8. Theo nghị định mới của Bộ GTVT về việc theo dõi hiện trường giao thông bằng Camera: " +
                "Bên B sử dụng vi phạm luật GTĐB, gây tai nạn cho người tham gia giao thông, bỏ trốn khỏi hiện trường," +
                " vì lý do thực tế chưa xử lý ngay được, sau một thời gian bị phát hiện hoặc Cơ quan Pháp luật điều tra " +
                "được thì vẫn phải chịu trách nhiệm trước Bên A và cơ quan pháp luật mặc dù hợp đồng kết thúc, Bên B đã trả xe.");

            //10
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("9. Mọi sự cố do Bên B gây ra, không thể tự giải quyết được phải nhờ đến Bên A trực tiếp giải quyết giúp, thì" +
                " tất cả chi phí đi lại, ăn, nghỉ của Bên A do Bên B thanh toán.");

            //11
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("10. Không sử dụng chất kích thích khi lái xe, nếu vi phạm bên B sẽ chịu toàn bộ trách nhiệm " +
                "trước pháp luật và sẽ chịu toàn bộ thiệt hại liên quan đến ngày nằm chờ gián đoạn kinh doanh, chi phí phát sinh nếu có.");

            #endregion

            #region Doan7
            //1
            para = doc.CreateParagraph();
            para.Alignment = ParagraphAlignment.LEFT;
            run = para.CreateRun();
            run.FontSize = 11;
            run.IsBold = true;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("Điều V: Các thỏa thuận đặt biệt ");

            //2
            para = doc.CreateParagraph();
            para.Alignment = ParagraphAlignment.LEFT;
            run = para.CreateRun();
            run.FontSize = 11;
            run.IsBold = true;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("1, Bên B cam kết không thực hiện các hành vi:");

            //3
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("a, Sử dụng xe thuê vào mục đích cầm cố hay thế chấp, sử dụng xe sai mục đích. Không được lái xe ra khỏi lãnh thổ Việt Nam.  ");

            //4
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("b, Sử dụng xe thuê vào những mục đích phi pháp như: vận chuyển hàng hóa trái phép (ma túy," +
                " chất cấm, hàng lậu, những đối tượng bị truy nã...)  ");

            //5
            para = doc.CreateParagraph();
            para.Alignment = ParagraphAlignment.LEFT;
            run = para.CreateRun();
            run.FontSize = 11;
            run.IsBold = true;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("2, Bên A có quyền:");

            //6
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("a, Báo cho cơ quan, gia đình Bên B, cơ quan điều tra nếu Bên B cố tình không liên lạc với Bên A; ");

            //7
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("b, Thanh lý tài sản thế chấp của Bên B, đền bù thiệt hại cho Bên A do bên B gây ra, nếu thiếu " +
                "Bên A sẽ tiếp tục truy thu hoặc thực hiện các biện pháp theo quy định của pháp luật nhằm đảm bảo quyền" +
                " lợi chính đáng của mình cho đến khi giải quyết xong thiệt hại; ");
            
            //8
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("c, Bên A có quyền đơn phương hủy hợp đồng nếu thấy lái xe Bên B hoặc Bên B không đảm bảo " +
                "kỹ thuật- chất lượng cho xe, cho an toàn giao thông; ");
            #endregion

            #region Doan8
            //1
            para = doc.CreateParagraph();
            para.Alignment = ParagraphAlignment.LEFT;
            run = para.CreateRun();
            run.FontSize = 11;
            run.IsBold = true;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("Điều VI: Cam kết chung:  ");

           

            //2
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("1. Hai bên cam kết thực hiện đúng các quy định trong hợp đồng. " +
                "Trong trường hợp xảy ra tranh chấp, hai bên chủ động cùng nhau thương lượng, giải" +
                " quyết. Nếu không thành công thì hai bên cùng đưa vụ việc ra tòa án có thẩm quyền giải quyết. ");

            //3
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("2. Hợp đồng này được lập thành 02 bản có giá trị pháp lý như nhau và có hiệu lực kể từ ngày ký. ");
            #endregion

            #region Doan9
            XWPFTable table = doc.CreateTable(1, 4);
            table.Width = 4000;
            table.GetRow(0).GetCell(0).SetText("Bên A");
            table.GetRow(0).GetCell(3).SetText("BênB");

            #endregion
            //GHI FILE
            MemoryStream ms = new MemoryStream();
            doc.Write(ms);
            var bytes = ms.ToArray();
            ms.Close();
            return File(bytes, "application/msword", "st.docx");
        }
    }
}