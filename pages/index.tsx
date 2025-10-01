import Head from "next/head";

export default function Home() {
  return (
    <>
      <Head>
        <title>Goobike Auto</title>
      </Head>
      <main style={{ padding: "2rem", fontFamily: "system-ui, sans-serif" }}>
        <h1>Goobike Auto 管理</h1>
        <p>このアプリケーションは Google Apps Script と連携するコードの管理リポジトリです。</p>
      </main>
    </>
  );
}
