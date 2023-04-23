import * as S from './styles'

export default function Banner({ title, image }) {
  return (
    <S.Banner
      style={{ backgroundImage: 'url(' + 'https://admin.parceirofacolchoes.com.br' + image + ')' }}
    >
      {title}
    </S.Banner>
  )
}
