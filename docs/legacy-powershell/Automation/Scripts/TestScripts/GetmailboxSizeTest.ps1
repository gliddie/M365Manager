Get-Mailbox "61177" -ResultSize Unlimited |

  Get-MailboxStatistics |

  Select DisplayName, `

  @{name=”TotalItemSize (GB)”; expression={[math]::Round( `

    ($_.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1GB),2)}}, `

  ItemCount
  
# |

#Sort “TotalItemSize (GB)” -Descending
