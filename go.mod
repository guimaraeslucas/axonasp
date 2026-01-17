module go-asp

go 1.24.0

toolchain go1.24.2

require (
	github.com/denisenkom/go-mssqldb v0.12.3
	github.com/guimaraeslucas/vbscript-go v0.9.0
	github.com/joho/godotenv v1.5.1
	golang.org/x/crypto v0.47.0
)

require (
	filippo.io/edwards25519 v1.1.0 // indirect
	github.com/dustin/go-humanize v1.0.1 // indirect
	github.com/google/uuid v1.6.0 // indirect
	github.com/mattn/go-isatty v0.0.20 // indirect
	github.com/ncruces/go-strftime v1.0.0 // indirect
	github.com/remyoudompheng/bigfft v0.0.0-20230129092748-24d4a6f8daec // indirect
	golang.org/x/exp v0.0.0-20251023183803-a4bb9ffd2546 // indirect
	golang.org/x/sys v0.40.0 // indirect
	gopkg.in/alexcesaro/quotedprintable.v3 v3.0.0-20150716171945-2caba252f4dc // indirect
	modernc.org/libc v1.67.6 // indirect
	modernc.org/mathutil v1.7.1 // indirect
	modernc.org/memory v1.11.0 // indirect
)

require (
	github.com/go-sql-driver/mysql v1.9.3
	github.com/golang-sql/civil v0.0.0-20190719163853-cb61b32ac6fe // indirect
	github.com/golang-sql/sqlexp v0.1.0 // indirect
	github.com/lib/pq v1.10.9
	gopkg.in/gomail.v2 v2.0.0-20160411212932-81ebce5c23df
	modernc.org/sqlite v1.44.1
)

replace github.com/guimaraeslucas/vbscript-go => ./VBScript-Go
