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
                ChangeUIStatus(Color.FromArgb(40, 80, 40), "Объекты KSC на данном компьютере поддерживаются.");
            }
            else
            {
                ChangeUIStatus(Color.FromArgb(80, 40, 40), "Объекты KSC на данном компьютере не поддерживаются. " +
                    "Для их поддержки установите Kaspersky Security Center.");

                Task.Run(() =>
                {
                    MessageBox.Show("Установите Kaspersky Security Center для поддержки KSC объектов.",
                        "Отсутствует поддержка KSC объектов",
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