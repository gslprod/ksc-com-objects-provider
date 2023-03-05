using TVCP_Module___COM_Objects_Provider.KSC;
using TVCP_Module___COM_Objects_Provider.NamedPipeConnection;

namespace TVCP_Module___COM_Objects_Provider
{
    public partial class MainForm : Form
    {

        public MainForm()
        {
            InitializeComponent();

            CheckForCompatibility();
            NamedPipeServer.StartConnectionThread();
        }

        private void ExitButton_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void CheckForCompatibility()
        {
            if (KSCProvider.CheckForObjectsAviability())
            {
                ChangeUIStatus(Color.FromArgb(40, 80, 40), "������� KSC �� ������ ���������� ��������������.");
            }
            else
            {
                ChangeUIStatus(Color.FromArgb(80, 40, 40), "������� KSC �� ������ ���������� �� ��������������. " +
                    "��� �� ��������� ���������� Kaspersky Security Center.");

                Task.Run(() =>
                {
                    MessageBox.Show("���������� Kaspersky Security Center ��� ��������� KSC ��������.",
                        "����������� ��������� KSC ��������",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                });
            }
        }

        private void ChangeUIStatus(Color background, string text)
        {
            stateLabel.Text = text;
            stateLabel.BackColor = background;
        }

        private void RefreshButton_Click(object sender, EventArgs e)
        {
            CheckForCompatibility();
        }
    }
}