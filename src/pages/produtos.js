import Menu from 'components/menu'
import Nav from 'components/nav'
import { useRequest } from 'hooks/requests'

export default function Produtos() {
  const { data: teste } = useRequest(`/produtos`)
  console.log(teste)
  return (
    <>
      <Nav />
      <Menu />
      <h1>Produtos</h1>
      {teste?.data?.map((item, i) => (
        <li key={'produtos' + i}>{item.attributes.name}</li>
      ))}
    </>
  )
}
