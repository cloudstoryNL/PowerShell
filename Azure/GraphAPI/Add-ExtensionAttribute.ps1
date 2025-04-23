$params = @{
  onPremisesExtensionAttributes = @{
    extensionAttribute14 = "corpnet.one-id.it/CORPNET/Germany/"
  }
}

Update-MgUser -UserId 'user@domain.com' -BodyParameter $params