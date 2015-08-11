#Office 365 Inbox Viewer for Windows 10#

I recently had a need to view an Office 365 email inbox for quick triage and task assignment as part of a larger project. This tutorial describes how to create a mail client for an Office 365 inbox.

Note: If you are using the sample code just to try out the application you must complete steps 5 through 9 below to create your application definition in Azure Active Directory and to ad the Azure AD ClientID into the App.xaml file.

## Create the Project and add references ##
1. Open Visual Studio 2015.
2. File | New | Project...
3. Select Visual C# | Windows | Universal | Blank App (Universal Windows)
4. Name the Project MailClientWin10App. Click **OK**
5. In Solution Explorer right-click on References then click Add Connected Service. Select **Office 365 APIs** then click **Configure**.
6. Enter you Office 365 developer domain (i.e. contosodev.onmicrosoft.com), click Next and enter your developer tenant credentials.
7. Select *Create new new Azure AD application to access Office 365 API services*, click Next.
8. Click Next until you arrive on Mail settings and select the *Read your mail* checkbox:
![Office 365 Mail Permissions](http://i.imgur.com/VGI2o8i.png)
9. Click Finish.

## Removed the Frame Rate Counter ##
Let's first remove the ugly frame rate counter that shows up by default when debugging an app.

1. Right-click on App.xaml and click View Code. Set `EnableFrameRateCounter = false;` (around line 46) or remove the line completely.

##Set up Office 365 Authentication##

Now that the Office 365 Service refernece has been addded to the project, we can start developing code to authenticate against Azure Active Directory and retrieve email information from the Office 365 API. In order to generate the access token from Azure Active Directory we will use the [WebAccountProvider](https://msdn.microsoft.com/en-us/library/windows/apps/windows.security.credentials.webaccountprovider.aspx) class. To do this we will create an AuthUtil class our XAML page will use for authentication and creation of the Outlook client.

1. Right-click on the project and create a new class (Add | Class) named AuthUtil.cs.
2. Add the following using statements to the class:

    ```csharp
    
    using Microsoft.Office365.OutlookServices;
    using System.Threading.Tasks;
    using Windows.Security.Authentication.Web.Core;
    using Windows.Security.Credentials;
    
    ```
3. Add the following method which retrieves the access token using `WebAccountProvider`

    ```csharp
    
    private static async Task<string> GetAccessToken()
    {
        string token = null;
    
        //first try to get the token silently
        WebAccountProvider aadAccountProvider = await WebAuthenticationCoreManager.FindAccountProviderAsync("https://login.windows.net");
        WebTokenRequest webTokenRequest = new WebTokenRequest(aadAccountProvider, String.Empty, App.Current.Resources["ida:ClientID"].ToString(), WebTokenRequestPromptType.Default);
        webTokenRequest.Properties.Add("authority", "https://login.windows.net");
        webTokenRequest.Properties.Add("resource", "https://outlook.office365.com/");
        WebTokenRequestResult webTokenRequestResult = await WebAuthenticationCoreManager.GetTokenSilentlyAsync(webTokenRequest);
        if (webTokenRequestResult.ResponseStatus == WebTokenRequestStatus.Success)
        {
            WebTokenResponse webTokenResponse = webTokenRequestResult.ResponseData[0];
            token = webTokenResponse.Token;
        }
        else if (webTokenRequestResult.ResponseStatus == WebTokenRequestStatus.UserInteractionRequired)
        {
            //get token through prompt
            webTokenRequest = new WebTokenRequest(aadAccountProvider, String.Empty, App.Current.Resources["ida:ClientID"].ToString(), WebTokenRequestPromptType.ForceAuthentication);
            webTokenRequest.Properties.Add("authority", "https://login.windows.net");
            webTokenRequest.Properties.Add("resource", "https://outlook.office365.com/");
            webTokenRequestResult = await WebAuthenticationCoreManager.RequestTokenAsync(webTokenRequest);
            if (webTokenRequestResult.ResponseStatus == WebTokenRequestStatus.Success)
            {
                WebTokenResponse webTokenResponse = webTokenRequestResult.ResponseData[0];
                token = webTokenResponse.Token;
            }
        }
    
        return token;
    }
    
    ```
4. Add a method to retrieve the Office 365 Outlook client

    ```csharp
    
        public static async Task<OutlookServicesClient> EnsureClient()
        {
            return new OutlookServicesClient(new Uri("https://outlook.office365.com/ews/odata"), async () => {
                return await GetAccessToken();
            });
        }
    
    ```

## Converters to improve display ##
In order to display the email dates and attachment icon in a user-friendly manner, we will create a couple of converters which will be used by the XAML.

Create a new folder in your project named `Converters`.
 
### NullableBoolVisibilityConverter ###

The `NullableBoolVisibilityConverter` class will convert from a nullable bool into a Visibility enumeration. This will allow us to determine whether the attachment icon will be visible within the email list.

Create a new class named `NullableBoolVisibilityConverter` within the Converters folder and insert the following code:

    ```csharp
    
        using System;
        using Windows.UI.Xaml;
        using Windows.UI.Xaml.Data;
    
        namespace MailClientWin10App.Converters
        {
            class NullableBoolVisibilityConverter : IValueConverter
            {
                public object Convert(object value, Type targetType, object parameter, string language)
                {
                    bool? b = value as bool?;
                    if (b == null || !b.HasValue || !b.Value)
                        return Visibility.Collapsed;
                    else
                        return Visibility.Visible;
                    //if (b.Value) return Visibility.Visible;
                    //else return Visibility.Collapsed;
                }
    
                public object ConvertBack(object value, Type targetType, object parameter, string language)
                {
                    throw new NotImplementedException();
                }
            }
        }
    
    ```

### EmailDateToStringConverter ###
The `EmailDateToStringConverter` will do basic date formatting to show dates in a more friendly manner.

Create a new class named `EmailDateToStringConverter` within the Converter folder and insert the following code:

    ```csharp
    
        using System;
        using Windows.UI.Xaml.Data;
    
        namespace Spark.Life.Helpers
        {
            class EmailDateToStringConverter : IValueConverter
            {
                public object Convert(object value, Type targetType, object parameter, string language)
                {
                    DateTimeOffset? dateVal = value as DateTimeOffset?;
                    if (dateVal == null || !dateVal.HasValue)
                        return value;
    
                    var myDate = dateVal.Value.ToLocalTime();
                    string retVal = string.Empty;
                    if (myDate.Date == DateTime.Today)
                    {
                        retVal = myDate.ToString("h:mm tt");
                    }
                    else if (myDate.Date > DateTime.Today.AddDays(-6))
                    {
                        retVal = myDate.ToString("ddd h:mm tt");
                    }
                    else if (myDate.Year == DateTime.Today.Year)
                    {
                        retVal = myDate.Date.ToString("ddd M/dd");
                    }
                    else
                    {
                        retVal = myDate.Date.ToString("ddd M/dd/yy");
                    }
    
                    return retVal;
                }
    
                public object ConvertBack(object value, Type targetType, object parameter, string language)
                {
                    throw new NotImplementedException();
                }
            }
        }
    
    
    ```

## Creating the main page ##

Modify the MainPage.xaml.cs file to test the connection to Office 365.
 
1. Open MainPage.xaml and change the display to *13.3" Desktop (1280 x 720) 100% scale*:
 ![](http://i.imgur.com/nbTWSTn.png)
2. We need to import the a few namespaces so that we can reference them within our XAML. Add the following within the `<Page` element: 

    ```xml
        xmlns:outlook="using:Microsoft.Office365.OutlookServices"
        xmlns:converters="using:MailClientWin10App.Converters"
    ```

3. Before the opening `<Grid` tag, we will add our converters and a template to render the list of emails. Add the following code:

    ```xml
    
        <Page.Resources>   
            <DataTemplate x:Key="MasterListViewItemTemplate" x:DataType="outlook:IMessage">
                <StackPanel Margin="12,11,12,13">
                    <Grid Grid.ColumnSpan="2">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*" />
                            <ColumnDefinition Width="Auto" />
                        </Grid.ColumnDefinitions>
    
                        <TextBlock Text="{x:Bind From.EmailAddress.Name}" MaxLines="1"
                               Foreground="{ThemeResource SystemControlForegroundBaseHighBrush}"
                               Style="{ThemeResource SubtitleTextBlockStyle}" />
    
                        <SymbolIcon Symbol="Attach" Grid.Column="1" Foreground="{ThemeResource SystemControlForegroundBaseMediumBrush}" Visibility="{x:Bind HasAttachments,Converter={StaticResource NullableBoolConverter}}" />
    
                    </Grid>
                    <TextBlock Text="{x:Bind Subject}" Grid.Row="1" MaxLines="1"
                               Foreground="{ThemeResource SystemControlForegroundBaseMediumBrush}"
                               Style="{ThemeResource BodyTextBlockStyle}" Grid.ColumnSpan="2" />
    
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*" />
                            <ColumnDefinition Width="Auto" />
                        </Grid.ColumnDefinitions>
    
                        <TextBlock Text="{x:Bind BodyPreview}" Style="{ThemeResource BodyTextBlockStyle}" 
                                Foreground="{ThemeResource SystemControlForegroundBaseMediumBrush}"
                                MaxLines="1" />
                        <TextBlock Text="{x:Bind DateTimeSent}" Grid.Column="1" Margin="12,2,0,0"
                               Foreground="{ThemeResource SystemControlForegroundBaseMediumBrush}"
                               Style="{ThemeResource BodyTextBlockStyle}" />
                    </Grid>
                </StackPanel>
            </DataTemplate>
    
        </Page.Resources>
    
    ```

3. Between the grid tags, we need to add our layout which will be a simple 2 row, 2 column grid:

    ```xml
    
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto" />
                <RowDefinition Height="*" />
            </Grid.RowDefinitions>
            <Grid.ColumnDefinitions>
                <ColumnDefinition x:Name="MasterColumn" Width="320" />
                <ColumnDefinition x:Name="DetailColumn" Width="*" />
            </Grid.ColumnDefinitions>
    
    ```

4. Add an inbox header and refresh button as another grid inside of the first grid and a progress indicator to display while the messages are being retrieved. Add the following just after the `</Grid.ColumnDefinitions>` tag:

    ```xml
    
            <Grid Background="{ThemeResource SystemControlBackgroundChromeMediumBrush}">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*" />
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Text="Inbox" Margin="24,8,8,8"
                           Style="{ThemeResource TitleTextBlockStyle}" />
                <AppBarButton Icon="Refresh" Grid.Column="1" x:Name="RefreshButton" Click="RefreshButton_Click" />
            </Grid>
    
            <ProgressRing x:Name="progressRing" HorizontalAlignment="Center" VerticalAlignment="Top" Margin="8,24,8,0" Grid.Row="1"/>
    
    ```

5. Add a list to display the inbox. We'll use a ListView along with the template we created earlier. Add the following after the `<ProgressRing />` element:

    ```xml
    
            <ListView x:Name="MasterListView" Grid.Row="1" ItemContainerTransitions="{x:Null}"
                      ItemTemplate="{StaticResource MasterListViewItemTemplate}"
                      IsItemClickEnabled="True"
                      ItemClick="MasterListView_ItemClick">
                <ListView.ItemContainerStyle>
                    <Style TargetType="ListViewItem">
                        <Setter Property="HorizontalContentAlignment" Value="Stretch" />
                        <Setter Property="BorderBrush" Value="{ThemeResource SystemControlForegroundBaseLowBrush}" />
                        <Setter Property="BorderThickness" Value="0,0,0,1" />
                    </Style>
                </ListView.ItemContainerStyle>
            </ListView>
    
    ```


6. Wire up code to retrieve emails. In Solution Explorer, Right-click MainPage.xaml and select View Code. Add a new method to load emails:

    ```csharp
    
            private async void LoadEmailMessagesFromOffice365()
            {
                MasterListView.ItemsSource = null;
                progressRing.IsActive = true;
                
                var outlookClient = await AuthUtil.EnsureClient();
    
                var messages = await outlookClient.Me.Folders["Inbox"].Messages.OrderByDescending(m => m.DateTimeReceived).Take(50).ExecuteAsync();
    
                progressRing.IsActive = false;
    
                MasterListView.ItemsSource = messages.CurrentPage;
            }
    
    ```

7. Create a method for the refresh button click event, call the load method when the page loads and a blank click handler for the list:

    ```chsarp
    
            private void RefreshButton_Click(object sender, RoutedEventArgs e)
            {
                LoadEmailMessagesFromOffice365();
            }
            protected override void OnNavigatedTo(NavigationEventArgs e)
            {
                base.OnNavigatedTo(e);
                LoadEmailMessagesFromOffice365();
            }
    
            private void MasterListView_ItemClick(object sender, ItemClickEventArgs e)
            {
    
            }
    
    ```
3. Run your project using Local Machine:

![Local Machine Debugging](http://i.imgur.com/skxKJ1Z.png)

The first time the project runs you will be prompted for your Office 365 username and password.

![Auth Dialog](http://i.imgur.com/lfYmGW7.png) 

Once authenticated you will be then prompted for consent to read your email.

![Consent Dialog](http://i.imgur.com/54m65ev.png)

The dialog will provide additional information to the user about what is going on. Then, if everything is successful, your email will be displayed:

![Inbox Start](http://i.imgur.com/K6mZaY6.png)

Click the refresh button and check to make sure the inbox re-loads.

## Showing the email details ##

Now that we have the inbox displaying properly, let's create XAML and code to display the full subject and body of the email, including any HTML formatting.

We will add another DataTemplate that contains a simple grid containing a stack panel for the header and a WebView for displaying the HTML. However, before we can use binding with the web view, we first need to add a control extension.

Right-click on the Converters folder and add a new class named ControlExtensions. Add the following code:

```csharp

    using Windows.UI.Xaml;
    using Windows.UI.Xaml.Controls;

    namespace MailClientWin10App.Converters
    {
        class ControlExtensions
        {

            public static string GetHTML(DependencyObject obj)
            {
                return (string)obj.GetValue(HTMLProperty);
            }

            public static void SetHTML(DependencyObject obj, string value)
            {
                obj.SetValue(HTMLProperty, value);
            }

            // Using a DependencyProperty as the backing store for HTML.  This enables animation, styling, binding, etc...  
            public static readonly DependencyProperty HTMLProperty =
                DependencyProperty.RegisterAttached("HTML", typeof(string), typeof(ControlExtensions), new PropertyMetadata(0, new PropertyChangedCallback(OnHTMLChanged)));

            private static void OnHTMLChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
            {
                WebView wv = d as WebView;
                if (wv != null)
                {
                    wv.NavigateToString((string)e.NewValue);
                }
            }


        }
    }


```

Now, let's get back to the XAML. Open MainPage.xaml again and insert the following template just before the `</Page.Resources>` tag.

``` xml

    <DataTemplate x:Key="DetailContentTemplate" x:DataType="outlook:IMessage">
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto" />
                <RowDefinition Height="*" />
            </Grid.RowDefinitions>

            <StackPanel HorizontalAlignment="Stretch">
                <TextBlock Text="{x:Bind From.EmailAddress.Name}" Grid.Row="1" MaxLines="1"
                       Foreground="{ThemeResource SystemControlForegroundBaseMediumBrush}"
                       Style="{ThemeResource SubtitleTextBlockStyle}" />
                <TextBlock Text="{x:Bind DateTimeSent,Converter={StaticResource EmailDateToStringConverter}}" Grid.Row="2" Grid.Column="1" Margin="12,2,0,0"
                       Foreground="{ThemeResource SystemControlForegroundBaseMediumBrush}"
                       Style="{ThemeResource BodyTextBlockStyle}" />
                <TextBlock Margin="0,8" Style="{ThemeResource TitleTextBlockStyle}"
                       HorizontalAlignment="Left" Text="{x:Bind Subject}"/>
            </StackPanel>

            
            <WebView x:Name="DetailContentWebView" converters:ControlExtensions.HTML="{x:Bind Body.Content}" 
                        Width="{Binding VisibleWidth, ElementName=DetailColumn,Mode=OneWay}" VerticalAlignment="Stretch"
                        ScrollViewer.VerticalScrollBarVisibility="Auto" ScrollViewer.VerticalScrollMode="Auto" Grid.Row="1"
                      />
            
        </Grid>
    </DataTemplate>


```

Now, let's add a content presenter control just before the final `</Grid>` tag on the page:

```xaml

    <ContentPresenter
        x:Name="DetailContentPresenter"
        Grid.Column="1"
        Grid.RowSpan="2"
        BorderThickness="1,0,0,0"
        Padding="24,0"
        BorderBrush="{ThemeResource SystemControlForegroundBaseLowBrush}"
        Content="{x:Bind MasterListView.SelectedItem, Mode=OneWay}"
        ContentTemplate="{StaticResource DetailContentTemplate}">
    </ContentPresenter>


```
## Contgratulations! ##

You just built a working Office 365 inbox viewer in Windows 10.

![Working Inbox Viewer](http://i.imgur.com/FzKa44c.png)