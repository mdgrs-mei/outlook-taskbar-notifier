Add-Type -TypeDefinition @"
public class DelegateCommand : System.Windows.Input.ICommand
{
    public DelegateCommand(System.Action<object> action)
    {
        m_action = action;
    }

    public bool CanExecute(object parameter)
    {
        return true;
    }

    public void Execute(object parameter)
    {
        m_action(parameter);
    }

    public event System.EventHandler CanExecuteChanged = delegate {};
    private System.Action<object> m_action;
}
"@
