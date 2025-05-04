import React, { useState } from 'react';
import { MsalProvider, useMsal, useIsAuthenticated } from '@azure/msal-react';
import { PublicClientApplication } from '@azure/msal-browser';
import { msalConfig, loginRequest } from './authConfig';
import {
  FluentProvider,
  webLightTheme,
  Card,
  CardHeader,
  Button,
  Title1,
  Title3,
  Body1,
  makeStyles,
  tokens,
  Avatar,
  Divider,
} from '@fluentui/react-components';
import {
  PersonRegular,
  SignOutRegular,
  CodeRegular,
  TvRegular,
  StorageRegular,
  BoxMultipleRegular,
  BoxSearchRegular,
} from '@fluentui/react-icons';
import { SharePointSites } from './components/SharePointSites';
import { Drives } from './components/Drives';
import { SharePointContainers } from './components/SharePointContainers';
import { ContainerTypes } from './components/ContainerTypes';
import './App.css';

// Create MSAL instance
const msalInstance = new PublicClientApplication(msalConfig);

const useStyles = makeStyles({
  root: {
    minHeight: '100vh' as unknown as never,
    backgroundColor: tokens.colorNeutralBackground2 as unknown as never,
    display: 'flex' as unknown as never,
    flexDirection: 'column' as unknown as never,
    alignItems: 'center' as unknown as never,
    paddingTop: tokens.spacingVerticalXXXL as unknown as never,
    paddingBottom: tokens.spacingVerticalXXXL as unknown as never,
    paddingLeft: tokens.spacingHorizontalXXL as unknown as never,
    paddingRight: tokens.spacingHorizontalXXL as unknown as never,
  },
  header: {
    display: 'flex' as unknown as never,
    flexDirection: 'column' as unknown as never,
    alignItems: 'center' as unknown as never,
    gap: tokens.spacingVerticalXL as unknown as never,
    marginBottom: tokens.spacingVerticalXXXL as unknown as never,
  },
  card: {
    width: '100%' as unknown as never,
    maxWidth: '600px' as unknown as never,
    padding: tokens.spacingHorizontalXL as unknown as never,
    boxSizing: 'border-box' as unknown as never,
  },
  userInfo: {
    display: 'flex' as unknown as never,
    alignItems: 'center' as unknown as never,
    gap: tokens.spacingHorizontalL as unknown as never,
    padding: tokens.spacingHorizontalL as unknown as never,
  },
  buttonsContainer: {
    display: 'flex' as unknown as never,
    flexDirection: 'column' as unknown as never,
    gap: tokens.spacingVerticalL as unknown as never,
    marginTop: tokens.spacingVerticalXL as unknown as never,
  },
  signInCard: {
    width: '100%' as unknown as never,
    maxWidth: '400px' as unknown as never,
    textAlign: 'center' as unknown as never,
  },
  signInContent: {
    marginTop: tokens.spacingVerticalXL as unknown as never,
    marginBottom: tokens.spacingVerticalXL as unknown as never,
  },
  icon: {
    fontSize: '48px' as unknown as never,
    color: tokens.colorBrandForeground1 as unknown as never,
  },
  mainContent: {
    width: '100%' as unknown as never,
    maxWidth: '1200px' as unknown as never,
    display: 'flex' as unknown as never,
    flexDirection: 'column' as unknown as never,
    alignItems: 'center' as unknown as never,
  },
});

type AppView = 'home' | 'sharepoint-sites' | 'drives' | 'sharepoint-containers' | 'container-types';

function SignInButton() {
  const { instance } = useMsal();
  const styles = useStyles();

  const handleLogin = () => {
    instance.loginPopup({
      ...loginRequest,
      prompt: "consent"  // Force consent prompt
    }).catch((e: any) => {
      console.error(e);
    });
  };

  return (
    <Card className={styles.signInCard}>
      <CardHeader
        header={<Title3>Sign in to Microsoft Graph Manager</Title3>}
      />
      <Body1 className={styles.signInContent}>
        Sign in with your Microsoft 365 account to manage SharePoint sites and drives.
      </Body1>
      <Button
        appearance="primary"
        icon={<PersonRegular />}
        onClick={handleLogin}
        size="large"
      >
        Sign in with Microsoft
      </Button>
    </Card>
  );
}

function SignOutButton() {
  const { instance } = useMsal();

  const handleLogout = () => {
    instance.logoutPopup().catch((e: any) => {
      console.error(e);
    });
  };

  return (
    <Button
      appearance="outline"
      icon={<SignOutRegular />}
      onClick={handleLogout}
    >
      Sign Out
    </Button>
  );
}

function WelcomeUser() {
  const { accounts } = useMsal();
  const account = accounts[0];
  const styles = useStyles();

  return (
    <div className={styles.userInfo}>
      <Avatar 
        name={account?.name || 'User'} 
        size={48}
      />
      <div>
        <Title3>{account?.name}</Title3>
        <Body1>{account?.username}</Body1>
      </div>
    </div>
  );
}

function AuthenticatedApp() {
  const isAuthenticated = useIsAuthenticated();
  const styles = useStyles();
  const [currentView, setCurrentView] = useState<AppView>('home');

  const navigateTo = (view: AppView) => {
    setCurrentView(view);
  };

  const renderContent = () => {
    switch (currentView) {
      case 'sharepoint-sites':
        return <SharePointSites onBack={() => navigateTo('home')} />;
      case 'drives':
        return <Drives onBack={() => navigateTo('home')} />;
      case 'sharepoint-containers':
        return <SharePointContainers onBack={() => navigateTo('home')} />;
      case 'container-types':
        return <ContainerTypes onBack={() => navigateTo('home')} />;
      default:
        return (
          <Card className={styles.card}>
            <WelcomeUser />
            <Divider />
            <div className={styles.buttonsContainer}>
              <Button
                appearance="primary"
                size="large"
                icon={<TvRegular />}
                onClick={() => navigateTo('sharepoint-sites')}
              >
                List SharePoint Sites
              </Button>
              <Button
                appearance="primary"
                size="large"
                icon={<StorageRegular />}
                onClick={() => navigateTo('drives')}
              >
                View Drives
              </Button>
              <Button
                appearance="primary"
                size="large"
                icon={<BoxMultipleRegular />}
                onClick={() => navigateTo('sharepoint-containers')}
              >
                List SharePoint Containers
              </Button>
              <Button
                appearance="primary"
                size="large"
                icon={<BoxSearchRegular />}
                onClick={() => navigateTo('container-types')}
              >
                View Container Types
              </Button>
              <SignOutButton />
            </div>
          </Card>
        );
    }
  };

  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <CodeRegular className={styles.icon} />
        <Title1>Microsoft Grap API Manager</Title1>
      </div>
      
      {isAuthenticated ? (
        <div className={styles.mainContent}>
          {renderContent()}
        </div>
      ) : (
        <SignInButton />
      )}
    </div>
  );
}

function App() {
  return (
    <FluentProvider theme={webLightTheme}>
      <MsalProvider instance={msalInstance}>
        <AuthenticatedApp />
      </MsalProvider>
    </FluentProvider>
  );
}

export default App;
