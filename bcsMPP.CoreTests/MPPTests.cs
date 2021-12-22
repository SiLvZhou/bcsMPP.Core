using NUnit.Framework;
using bcsMPP.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aras.IOM;
using Aras;
using Aras.Server;

namespace bcsMPP.Core.Tests
{
    [TestFixture()]
    public class MPPTests
    {
        [Test()]
        public void getProcessPlanStructureTest()
        {
            HttpServerConnection connection = GetServerConnection();

            Innovator inn = IomFactory.CreateInnovator(connection);

            Item requestItem = inn.newItem("mpp_ProcessPlan", "mpp_getProcessPlanStructure");
            requestItem.setProperty("plan_ids", "");
            requestItem.setProperty("with_details_ids", "3874CB66E295400787701F20389666B9");//mpp_ProcessPlan(工艺流程规划)某一条数据的ID
            requestItem.setProperty("lang_code", "");
            requestItem.setProperty("bcs_location", "");

            bcsMPP.Core.MPP bcsMPP = new bcsMPP.Core.MPP(inn);
            Item result = bcsMPP.getProcessPlanStructure(requestItem);

            connection.Logout();

            Assert.Pass("");
        }

        [Test()]
        public void GetTestWorkHourTreeGridTest()
        {
            HttpServerConnection connection = GetServerConnection();

            Innovator inn = IomFactory.CreateInnovator(connection);

            Item requestItem = inn.newItem("mpp_ProcessPlan");
            requestItem.setProperty("processplan_id", "3874CB66E295400787701F20389666B9");//mpp_ProcessPlan(工艺流程规划)某一条数据的ID
            requestItem.setProperty("processplan_number", "11111");
            requestItem.setProperty("processplan_name", "22222");
            requestItem.setProperty("location_id", "");

            bcsMPP.Core.MPP bcsMPP = new bcsMPP.Core.MPP(inn);
            Item result = bcsMPP.GetTestWorkHourTreeGrid(requestItem);

            connection.Logout();
        }


        public HttpServerConnection GetServerConnection()
        {
            HttpServerConnection connection = IomFactory.CreateHttpServerConnection("http://localhost/InnMPPV2", "innChe1", "admin", "innovator");
            Item loginResult = connection.Login();
            if (loginResult.isError())
            {
                Assert.Fail("登录失败");
            }
            return connection;
        }

        [Test()]
        public void updateRelationPQDTest()
        {
            HttpServerConnection connection = GetServerConnection();

            Innovator inn = IomFactory.CreateInnovator(connection);

            Item mppItem = inn.getItemById("mpp_ProcessPlan", "3874CB66E295400787701F20389666B9");//mpp_ProcessPlan(工艺流程规划)某一条数据的ID

            bcsMPP.Core.MPP bcsMPP = new bcsMPP.Core.MPP(inn);
            Item result = bcsMPP.updateRelationPQD(mppItem);

            connection.Logout();

            Assert.Pass("");
        }

        [Test()]
        public void PBOM2MBOMTest()
        {
            HttpServerConnection connection = GetServerConnection();

            Innovator inn = IomFactory.CreateInnovator(connection);

            Item requestItem = inn.newItem("mpp_ProcessPlan");
            requestItem.setID("3874CB66E295400787701F20389666B9");//mpp_ProcessPlan(工艺流程规划)某一条数据的ID
            requestItem.setProperty("location", "");
            /// requestItem.setPropertyItem("rowitem", );
            requestItem.setProperty("errorString", "");

            bcsMPP.Core.MPP bcsMPP = new bcsMPP.Core.MPP(inn);
            Item result = bcsMPP.PBOM2MBOM(requestItem);

            connection.Logout();

        }
        StringBuilder gridStyle = new StringBuilder();
        [Test()]
        public void PBOM2MBOMCheckTest()
        {
            HttpServerConnection connection = GetServerConnection();
            Innovator inn = IomFactory.CreateInnovator(connection);
            //Aras.Server.Core.IOMConnection x = connection as Aras.Server.Core.IOMConnection;
            //Aras.Server.Core.CallContext CCO = x.CCO;
            //Aras.Server.Core.IContextState RequestState = CCO.RequestState;
            getGridHeader();
            Item requestItem = inn.newItem("mpp_ProcessPlan");
            requestItem.setID("3874CB66E295400787701F20389666B9");//mpp_ProcessPlan(工艺流程规划)某一条数据的ID
            requestItem.setProperty("location", "833FD89CEBC94E1EAEBDB1F4FD336363");
            requestItem.setProperty("rowitem", "");
            requestItem.setProperty("gridheaderxml", gridStyle.ToString());
            bcsMPP.Core.MPP bcsMPP = new bcsMPP.Core.MPP(inn);
            Item result = bcsMPP.PBOM2MBOMCheck(requestItem);

            connection.Logout();


        }
        private void getGridHeader()
        {
            gridStyle.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            gridStyle.Append("<table");
            gridStyle.Append(" font=\"Microsoft Sans Serif-8\"");
            gridStyle.Append(" sel_bgColor=\"steelbue\"");
            gridStyle.Append(" sel_TextColor=\"white\"");
            gridStyle.Append(" header_BgColor=\"buttonface\"");
            gridStyle.Append(" expandroot=\"true\"");
            gridStyle.Append(" expandall=\"false\"");
            gridStyle.Append(" treelines=\"1\"");
            gridStyle.Append(" editable=\"true\"");
            gridStyle.Append(" draw_grid=\"true\"");
            gridStyle.Append(" multiselect=\"false\"");
            gridStyle.Append(" column_draggable=\"true\"");
            gridStyle.Append(" enableHtml=\"false\"");
            gridStyle.Append(" enterAsTab=\"false\"");
            gridStyle.Append(" bgInvert=\"true\"");
            gridStyle.Append(" xmlns:msxsl=\"urn:schemas-microsoft-com:xslt\"");
            gridStyle.Append(" xmlns:aras=\"http://www.aras.com\"");
            gridStyle.Append(" xmlns:usr=\"urn:the-xml-files:xslt\">");

            //处理网格列
            gridStyle.Append("<thead>");
            gridStyle.Append("<th align=\"c\">物料信息</th>");
            gridStyle.Append("<th align=\"c\">检查结果</th>");
            gridStyle.Append("<th align=\"c\">物料版本</th>");
            gridStyle.Append("<th align=\"c\">物料状态</th>");
            gridStyle.Append("<th align=\"c\">数量</th>");
            gridStyle.Append("<th align=\"c\">BOM序号</th>");
            gridStyle.Append("<th align=\"c\">MPart</th>");
            gridStyle.Append("<th align=\"c\">错误信息</th>");
            gridStyle.Append("<th align=\"c\">执行动作</th>");
            gridStyle.Append("<th align=\"c\">同步</th>");
            gridStyle.Append("</thead>");

            gridStyle.Append("<columns>");
            gridStyle.Append("<column width=\"400\" edit=\"NOEDIT\" align=\"l\" order=\"1\" />");
            gridStyle.Append("<column width=\"60\" edit=\"NOEDIT\" align=\"c\" order=\"0\" />");
            gridStyle.Append("<column width=\"60\" edit=\"NOEDIT\" align=\"c\" order=\"2\" />");
            gridStyle.Append("<column width=\"70\" edit=\"NOEDIT\" align=\"c\" order=\"3\" />");
            gridStyle.Append("<column width=\"60\" edit=\"NOEDIT\" align=\"l\" order=\"4\" />");
            gridStyle.Append("<column width=\"60\" edit=\"NOEDIT\" align=\"c\" order=\"5\" />");
            gridStyle.Append("<column width=\"100\" edit=\"NOEDIT\" align=\"l\" order=\"6\" />");
            gridStyle.Append("<column width=\"200\" edit=\"NOEDIT\" align=\"l\" order=\"7\" />");
            gridStyle.Append("<column width=\"70\" edit=\"NOEDIT\" align=\"c\" order=\"8\" />");
            gridStyle.Append("<column width=\"40\" edit=\"FIELD\" align=\"c\" order=\"9\" />");
            gridStyle.Append("</columns>");
        }

        [Test()]
        public void mppExport2pdfTest()
        {
            HttpServerConnection connection = GetServerConnection();
            Innovator inn = IomFactory.CreateInnovator(connection);
            //Aras.Server.Core.IOMConnection x = connection as Aras.Server.Core.IOMConnection;
            //Aras.Server.Core.CallContext CCO = x.CCO;
            //Aras.Server.Core.IContextState RequestState = CCO.RequestState;
            Item requestItem = inn.newItem("mpp_ProcessPlan");
            requestItem.setProperty("html", "");
            bcsMPP.Core.MPP bcsMPP = new bcsMPP.Core.MPP(inn);
            Item result = bcsMPP.mppExport2pdf(requestItem, (Aras.Server.Core.IOMConnection)inn.getConnection());

            connection.Logout();

            Assert.Fail();
        }

        [Test()]
        public void readMPPTemplateFileTest()
        {
            HttpServerConnection connection = GetServerConnection();
            Innovator inn = IomFactory.CreateInnovator(connection);
            //Aras.Server.Core.IOMConnection x = connection as Aras.Server.Core.IOMConnection;
            //Aras.Server.Core.CallContext CCO = x.CCO;
            //Aras.Server.Core.IContextState RequestState = CCO.RequestState;
            Item requestItem = inn.newItem("mpp_ProcessPlan");
            requestItem.setProperty("html", "");
            bcsMPP.Core.MPP bcsMPP = new bcsMPP.Core.MPP(inn);
            Item result = bcsMPP.readMPPTemplateFile(requestItem,(Aras.Server.Core.IOMConnection)inn.getConnection());

            connection.Logout();

            Assert.Fail();
        }

        [Test()]
        public void mppProcessFlowUpdateTest()
        {
            HttpServerConnection connection = GetServerConnection();
            Innovator inn = IomFactory.CreateInnovator(connection);
            //Aras.Server.Core.IOMConnection x = connection as Aras.Server.Core.IOMConnection;
            //Aras.Server.Core.CallContext CCO = x.CCO;
            //Aras.Server.Core.IContextState RequestState = CCO.RequestState;
            Item requestItem = inn.newItem("mpp_ProcessPlan");
            requestItem.setID("3874CB66E295400787701F20389666B9");//旧版ID
            requestItem.setProperty("config_id", "3874CB66E295400787701F20389666B9");//和旧版ID对应的config_id
            bcsMPP.Core.MPP bcsMPP = new bcsMPP.Core.MPP(inn);
            Item result = bcsMPP.mppProcessFlowUpdate(requestItem);

            connection.Logout();

            //Assert.Fail();
        }
    }
}