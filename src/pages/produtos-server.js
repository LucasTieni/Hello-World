import axios from 'axios'
import Menu from 'components/menu'
import Nav from 'components/nav'

export default function Produtos({ produtos }) {
  return (
    <>
      <Nav />
      <Menu />
      <h1>Produtos</h1>
      {produtos?.data?.map((item, i) => (
        <li key={'produtos' + i}>{item.attributes.name}</li>
      ))}
    </>
  )
}

export async function getServerSideProps() {
  const produtosResponse = axios.get(process.env.NEXT_PUBLIC_STRAPI_API_URL + '/produtos')
  const [produtos] = await Promise.all([produtosResponse])

  return {
    props: { produtos: produtos.data },
  }
}
