package com.thopham.projects.dknec.demo_v2

import org.springframework.beans.factory.annotation.Autowired
import org.springframework.beans.factory.annotation.Value
import org.springframework.stereotype.Component
import javax.persistence.EntityManager

@Component
class TagRepository(val entityManager: EntityManager) {
    @Value("\${spring.profiles.active:nau_men}")
    lateinit var profile: String
    fun findAllByCreatedGreaterThanOrderByCreated(created: String): List<Tag> {
        val tagsName = findAllTagName()
        val hasAnyData = entityManager.createNativeQuery("select top 1 tag_value, created from z_tag_${tagsName[0]} where created > '$created'").resultList.filterNotNull().isNotEmpty()
        if(hasAnyData) {
            val tags = tagsName.map { tagName ->
                entityManager.createNativeQuery("select top 1 tag_value, created from z_tag_$tagName where created > '$created' order by created asc")
                        .resultList
                        .filterNotNull()
                        .map { tag -> Tag.fromSQL(tagName, tag) }
            }
                    .flatten()
                    .sortedByDescending { tag -> tag.created }
            val currentTime = tags.first().created
            return tags.map { tag ->
                tag.copy(created = currentTime)
            }
        }
        println("There are no new data...")
        return emptyList()
    }

    fun findAllOrderByCreatedAsc(): List<Tag> {
        val tagsName = findAllTagName()
        val tags = tagsName.map { tagName ->
            entityManager.createNativeQuery("select top 1 tag_value, created from z_tag_$tagName order by created asc")
                    .resultList
                    .filterNotNull()
                    .map { tag -> Tag.fromSQL(tagName, tag) }
        }
                .flatten()
                .sortedByDescending { tag -> tag.created }
        val currentTime = tags.first().created
        return tags.map { tag ->
            tag.copy(created = currentTime)
        }
    }

    fun findAllTagName(): Array<String> {
        val defaultTag = Constants.NAU_MEN_TAGS
        return if (profile == "nau_men")
            Constants.NAU_MEN_TAGS
        else if (profile == "chiet") Constants.CHIET_TAGS
        else defaultTag
    }
}