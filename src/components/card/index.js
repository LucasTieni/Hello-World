import * as S from './styles'
import Link from 'next/link'

export default function Card({ thumb, image, title, date, description, files, list }) {
  console.log('thumb', thumb)

  return (
    <S.Card>
      {thumb && (
        <S.Thumb
          style={{
            backgroundImage: 'url(' + 'https://admin.parceirofacolchoes.com.br' + thumb + ')',
          }}
        >
          {/* <img src={process.env.NEXT_PUBLIC_STRAPI + thumb} /> */}
          {/* <S.Empty>
            <img src={'/image-line.svg'} />
          </S.Empty> */}
        </S.Thumb>
      )}
      {title && <S.Title>{title}</S.Title>}
      {date && <S.Date>{date}</S.Date>}
      {description && <S.Description>{description}</S.Description>}
      {files &&
        files.map((file, i) => (
          <S.Links key={file + i}>
            <Link href={file.link} target="_blank">
              <img src={'/download-2-line.svg'} alt={file.tipo} /> {file.tipo}
            </Link>
          </S.Links>
        ))}
      {list &&
        list.map((item, i) => (
          <S.Links key={item + i}>
            {/* <Link href={item.link} target="_blank"> */}
            {item}
            {/* </Link> */}
          </S.Links>
        ))}
    </S.Card>
  )
}
